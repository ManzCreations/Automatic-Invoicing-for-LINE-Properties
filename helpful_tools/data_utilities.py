import math
import re

import pandas as pd


def reformat_and_update_files(df, customer_df, listing_column='Listing', customer_column='ListingBNB', reformat=True):
    """
    Reformat listings in a DataFrame and update customer codes based on another DataFrame.
    """
    replace_regex = r"[^A-Za-z0-9_]+"
    for index, row in df.iterrows():
        listing = row[listing_column]
        if pd.notnull(listing) and reformat:
            listing = re.sub(replace_regex, ' ', listing).strip()
            customer_info = customer_df[customer_df[customer_column] == listing]
            if not customer_info.empty:
                df.at[index, 'Code'] = customer_info['Code'].iloc[0]
                df.at[index, 'Customer'] = customer_info['QBO'].iloc[0]
    return df


def clean_dataframes(dataframes):
    """
    Clean specific dataframes by removing undesired rows.
    """
    for key, df in dataframes.items():
        if 'bnb' in key:
            df = df[~df['Type'].isin(['Resolution Payout', 'Resolution Adjustment'])]
            dataframes[key] = df
    return dataframes


def initialize_dataframes():
    """
    Initialize empty DataFrames for processing.
    """
    dataframes = {
        'bnb': pd.DataFrame(),
        'cleaning': pd.DataFrame(),
        'customer_info': pd.DataFrame(),
        'check': pd.DataFrame(),
        'vrbo': pd.DataFrame(),
        'entry_NCM': pd.DataFrame(),
        'entry_CM': pd.DataFrame(),
        'sales_entry': pd.DataFrame(),
        'credit_memo': pd.DataFrame(),
        'sales_receipts': pd.DataFrame(),
        'journal_entries': pd.DataFrame(),
        'checks': pd.DataFrame(),
        'tax_issues': pd.DataFrame(),
        'man_issues': pd.DataFrame(),
        'prev_cleaning': pd.DataFrame(),
    }
    return dataframes


def is_nan(value):
    """
    Check if a value is NaN.

    Args:
        value: The value to check.

    Returns:
        bool: True if the value is NaN, False otherwise.
    """
    try:
        return math.isnan(float(value))
    except ValueError:
        return False


def prepare_dataframe_columns(df, columns, replace_re, null_replacement='NULL'):
    """
    Cleans and prepares specified columns of a DataFrame.

    Args:
        df (pd.DataFrame): The DataFrame to process.
        columns (list): List of columns to include in the output DataFrame.
        replace_re (str): Regular expression pattern for replacement.
        null_replacement (str): Value to replace null/NaN entries.

    Returns:
        pd.DataFrame: The cleaned and prepared DataFrame.
    """
    df_col = pd.DataFrame(df, columns=columns)
    for col in columns:
        df_col.loc[df_col[col].isnull(), col] = null_replacement
        # Assuming the unidecode operation was handled elsewhere if needed
        df_col[col] = df_col[col].str.replace(replace_re, ' ', regex=True).str.strip()
    return df_col


def remove_extra_spaces(df, columns):
    """
    Removes extra spaces from specified columns in a DataFrame.

    Args:
        df (pd.DataFrame): The DataFrame to process.
        columns (list): List of column names to clean.

    Returns:
        pd.DataFrame: The DataFrame with cleaned columns.
    """
    for column in columns:
        df[column] = df[column].str.replace('  +', ' ', regex=True)
    return df


def find_diff_and_concat(source_df, target_df, source_column, target_column):
    """
    Finds differences between two columns in two DataFrames and concatenates the differing rows.

    Args:
        source_df (pd.DataFrame): DataFrame containing the source column.
        target_df (pd.DataFrame): DataFrame containing the target column.
        source_column (str): The source column name.
        target_column (str): The target column name.

    Returns:
        pd.DataFrame: Concatenated DataFrame of differences.
    """
    diff = set(source_df[source_column]) - set(target_df[target_column])
    diff_df = pd.DataFrame()

    for item in diff:
        new_rows = source_df[source_df[source_column] == item]
        diff_df = pd.concat([diff_df, new_rows], ignore_index=True)

    return diff_df


def set_dataframe_columns(df, columns):
    """
    Sets the column names of a DataFrame.
    """
    df.columns = columns
    return df


def convert_column_types(df, columns, dtype):
    """
    Converts the data types of specified columns in a DataFrame.
    """
    for column in columns:
        df[column] = df[column].astype(dtype)
    return df
