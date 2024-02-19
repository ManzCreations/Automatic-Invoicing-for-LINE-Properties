import datetime
import os
import re


def date_from_airbnb_name(file_name):
    """
    Determine month and year based on Airbnb filename.

    Args:
        file_name (str): Partial or full name of the Airbnb file.

    Returns:
        tuple: A tuple containing the month and year extracted from the file name.
    """
    for file in os.listdir('ModelFiles/.'):
        if fnmatch.fnmatch(file, '*.xlsx') and file_name in file:
            date_part = ''.join([char for char in file.split('-')[1] if char.isdigit()])
            month = int(date_part[:2])
            year = int(date_part[2:])
            return month, year
    return None, None


def read_reference_numbers(file_path='ref_number_values.txt'):
    """
    Read reference numbers from a file.

    Args:
        file_path (str, optional): Path to the reference numbers file. Defaults to 'ref_number_values.txt'.

    Returns:
        tuple: A tuple containing invoice, checks, and journal numbers.
    """
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()
            invoice_number = int(float(re.findall(r'\d+\.?\d*', lines[2])[0]))
            checks_number = int(float(re.findall(r'\d+\.?\d*', lines[3])[0]))
            journal_number = int(float(re.findall(r'\d+\.?\d*', lines[4])[0]))
    except FileNotFoundError:
        print(f'{file_path} not found! Please input starting values for reference numbers.\n')
        invoice_number = int(input('Starting reference number for Invoices: '))
        checks_number = int(input('Starting reference number for Checks: '))
        journal_number = int(input('Starting reference number for Journal Entries: '))

    return invoice_number, checks_number, journal_number


def write_reference_numbers(invoice_no, check_no, journal_no, month, year, file_path='ref_number_values.txt'):
    """
    Write reference numbers to a file.

    Args:
        invoice_no (int): Invoice number.
        check_no (int): Check number.
        journal_no (int): Journal number.
        month (int): Month.
        year (int): Year.
        file_path (str, optional): Path to the reference numbers file. Defaults to 'ref_number_values.txt'.
    """
    with open(file_path, 'w') as file:
        file.write(f"LIN date: {month}/{year}\n\n")
        file.write(f"Reference number (Invoice): {invoice_no}\n")
        file.write(f"Reference number (Checks): ABB TR {check_no:05d}\n")
        file.write(f"Reference number (Journal): PMT {journal_no:05d}\n")


def explanation_for_missing_data(df, explanation):
    """
    Add an explanation for missing data in a DataFrame.

    Args:
        df (pd.DataFrame): The DataFrame to modify.
        explanation (str): The explanation to add.

    Returns:
        pd.DataFrame: The modified DataFrame.
    """
    df.loc[0, 'Explanation'] = explanation
    return df


def month_number_to_name(month_num):
    """
    Convert a month number to its full name.

    Args:
        month_num (int): The month number.

    Returns:
        str: The full month name.
    """
    return datetime.datetime.strptime(str(month_num), "%m").strftime("%B")


def month_name_to_number(month_name):
    """
    Convert a month name to its number.

    Args:
        month_name (str): The full or abbreviated month name.

    Returns:
        int: The month number.
    """
    month_name = month_name.strip()[:3].lower()
    month_numbers = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                     'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
    return month_numbers[month_name]


def generate_dates(month, year):
    """
    Generate invoice and due dates based on month and year.

    Args:
        month (int): Month for the report.
        year (int): Year for the report.
    """
    next_month, next_year = (month % 12 + 1, year + (month // 12))
    invoice_date = datetime.datetime(next_year, next_month, 1)
    due_date = datetime.datetime(next_year, next_month, 5)

    return invoice_date.strftime('%m/%d/%Y'), due_date.strftime('%m/%d/%Y')
