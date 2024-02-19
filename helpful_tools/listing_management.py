import re
from collections import defaultdict

import pandas as pd


def separate_listings_based_on_output(cleaning_col, bnb, vrbo, path, month, year):
    """
    Separates listings into different Excel files based on the 'Output' column in the cleaning_col DataFrame.

    Args:
        cleaning_col (pd.DataFrame): DataFrame with cleaning information including 'Output', 'ListingBNB', and 'VRBO_ID' columns.
        bnb (pd.DataFrame): DataFrame with Airbnb listings.
        vrbo (pd.DataFrame): DataFrame with VRBO listings.
        path (str): Base path to save the Excel files.
        month (int): Current month for file naming.
        year (int): Current year for file naming.
    """
    aviad_excel = defaultdict(lambda: {'listings2find': [], 'vrbo_ids': []})

    for index, row in cleaning_col.iterrows():
        try:
            if isinstance(row['Output'], str) and re.match(r'(\w+)', row['Output']):
                aviad_excel[row['Output']]['listings2find'].append(row['ListingBNB'])
                aviad_excel[row['Output']]['vrbo_ids'].append(row['VRBO_ID'])
        except TypeError:
            pass  # Handle empty or NaN values

    for name, info in aviad_excel.items():
        filtered_bnb = bnb[bnb['Listing'].isin(info['listings2find'])]
        filtered_vrbo = vrbo[vrbo['Property ID'].isin(info['vrbo_ids'])]

        with pd.ExcelWriter(f"{path}\\{name} Reservations {month}_{year}.xlsx", engine='xlsxwriter') as writer:
            filtered_bnb.to_excel(writer, index=False, sheet_name='AirBNB')
            filtered_vrbo.to_excel(writer, index=False, sheet_name='VRBO')
