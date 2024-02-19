import os
from os.path import exists
from shutil import copyfile

import pandas as pd


def manage_vrbo_data(vrbo, vrbo_save_path, month, month_name, year, mtn, number2month):
    """
    Manages VRBO data, including filtering, updating, and saving to Excel.

    Args:
        vrbo (pd.DataFrame): The VRBO DataFrame to process.
        vrbo_save_path (str): Path where the VRBO data is saved.
        month (int): The current month as an integer.
        month_name (str): The current month as a name.
        year (int): The current year.
        mtn (function): Function to convert month name to number.
        number2month (function): Function to convert number to month name.
    """
    vrbo_wrong_month = vrbo[vrbo['Check-out'].dt.month != month]
    keep_locations = []
    vrbo_new = pd.DataFrame()

    if exists(vrbo_save_path):
        vrbo_new = pd.read_excel(vrbo_save_path)
        for idx, row in vrbo_wrong_month.iterrows():
            if not (vrbo_new['Reservation ID'] == row['Reservation ID']).any():
                vrbo_new = pd.concat([vrbo_new, pd.DataFrame([row])], ignore_index=True)
                vrbo_new.iloc[-1, 'Month'] = number2month(row['Check-out'].month)
                vrbo_new.iloc[-1, 'Year'] = row['Check-out'].year

        # Filter to keep relevant listings
        for idx, row in vrbo_new.iterrows():
            if mtn(row['Month']) >= month and row['Year'] >= year:
                keep_locations.append(idx)

        vrbo_new = vrbo_new.iloc[keep_locations]
        vrbo_new.loc[vrbo_new['Payout'] != 0, 'Payout'] = 0

        # Save updated VRBO data
        if not vrbo_new.empty:
            os.remove(vrbo_save_path)  # Ensure the file is removed before saving if it exists
            vrbo_new.to_excel(vrbo_save_path, index=False)


def copy_excel_file(source_path, destination_path):
    """
    Copies an Excel file from source to destination.

    Args:
        source_path (str): Source file path.
        destination_path (str): Destination file path.
    """
    copyfile(source_path, destination_path)


def auto_adjust_columns_width(writer, sheet_name, df):
    """
    Auto-adjusts column widths in an Excel sheet.
    """
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df):
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
        )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width


def write_dataframe_to_excel(writer, df, sheet_name, format_columns=None):
    """
    Writes a DataFrame to an Excel sheet and optionally applies formatting.
    """
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    if format_columns:
        workbook = writer.book
        format_col = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})
        worksheet = writer.sheets[sheet_name]
        for col_idx in format_columns:
            worksheet.set_column(col_idx, col_idx, None, format_col)
    auto_adjust_columns_width(writer, sheet_name, df)
