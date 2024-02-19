# customer_data_aggregation.py
import numpy as np
import pandas as pd


def isnan(value):
    """Check if a value is NaN."""
    try:
        return np.isnan(float(value))
    except (ValueError, TypeError):
        return False


def aggregate_customer_data(customer_col, cleaning_col, bnb_col, check_col, vrbo_col, month):
    """
    Aggregates data across multiple dataframes to create a comprehensive customer unit dataframe.

    Args:
        customer_col (pd.DataFrame): DataFrame containing customer-QBO information.
        cleaning_col (pd.DataFrame): DataFrame containing cleaning listings and related information.
        bnb_col (pd.DataFrame): DataFrame containing Airbnb listing information.
        check_col (pd.DataFrame): DataFrame containing check listing information.
        vrbo_col (pd.DataFrame): DataFrame containing VRBO listing information.
        month (int): Current month for processing.

    Returns:
        pd.DataFrame: Aggregated unit dataframe with comprehensive customer data.
    """
    unit = pd.DataFrame()
    for customer in customer_col['Customer-QBO']:
        customer_listing = cleaning_col[cleaning_col['QBO'] == customer]
        for index, row in customer_listing.iterrows():
            vrbo_90_day_amount = bnb_90_day_amount = income = vrbo_payout = vrbo_nights = 0
            vrbo_id = 'none'
            listing = bnb_col[bnb_col['Listing'] == row['ListingBNB']]
            if not listing.empty:
                # Logic for BNB amount and reservation count
                income, bnb_90_day_amount = calculate_bnb_income(listing, month)

            reservation_count = len(check_col[check_col['Listing'] == row['ListingBNB']])
            vrbo_payout, vrbo_nights, vrbo_90_day_amount = calculate_vrbo_payout(row['VRBO_ID'], vrbo_col, month)

            unit_new = pd.DataFrame({
                'Customer': [customer],
                'Listing': [row['ListingBNB']],
                'Income': [income],
                'CleaningFee': [row['Cleaning']],
                'Checkouts': [reservation_count],
                'TaxLocation': [row['Tax_Location']],
                # Add other columns as per the original structure
                'VRBO_ID': [vrbo_id],
                'VRBO_PAYOUT': [vrbo_payout],
                'VRBO_Nights': [vrbo_nights],
                'BNB_90_Day': [bnb_90_day_amount],
                'VRBO_90_Day': [vrbo_90_day_amount]
                # Continue adding other fields
            })
            unit = pd.concat([unit, unit_new], ignore_index=True)

    unit.columns = ['Customer', 'Listing', 'Income', 'CleaningFee', 'Checkouts', 'TaxLocation',
                    'VRBO_ID', 'VRBO_PAYOUT', 'VRBO_Nights', 'BNB_90_Day', 'VRBO_90_Day']
    return unit


def calculate_bnb_income(listing, month):
    """Calculates total income and 90 day term amounts for Airbnb listings."""
    income = listing['Amount'].sum()
    bnb_90_day_amount = listing[listing['Nights'] > 89]['Amount'].sum()
    return income, bnb_90_day_amount


def calculate_vrbo_payout(vrbo_id, vrbo_col, month):
    """Calculates total payout, nights, and 90 day term amounts for VRBO listings."""
    if isnan(vrbo_id):
        return 0, 0, 0

    vrbo_payouts = vrbo_col[vrbo_col['Property ID'] == vrbo_id]
    vrbo_payout = vrbo_payouts['Payout'].sum()
    vrbo_nights = vrbo_payouts['Nights'].sum()
    vrbo_90_day_amount = vrbo_payouts[vrbo_payouts['Nights'] > 89]['Payout'].sum()
    return vrbo_payout, vrbo_nights, vrbo_90_day_amount


def sort_and_prepare_unit(unit):
    """
    Sorts the unit DataFrame by customer and prepares it by filling NaN values.
    """
    unit = unit.sort_values(by=['Customer'])
    unit['CleaningFee'] = unit['CleaningFee'].apply(pd.to_numeric, errors='coerce').fillna(0)
    unit.loc[unit["CreditMemo"].isnull(), 'CreditMemo'] = 'NULL'
    return unit


def group_unit_by_credit_memo(unit):
    """
    Groups the unit DataFrame by the 'CreditMemo' column.
    """
    return unit.groupby('CreditMemo')
