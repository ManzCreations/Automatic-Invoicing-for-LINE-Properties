#######################################################################################################################
# Modules
import pip
import calendar
import datetime
import math
import os
import re
import fnmatch
import pandas as pd
import xlsxwriter
import numpy as np
import shutil
from shutil import copyfile
from calendar import month_name as mn
from os.path import exists
import openpyxl
from collections import defaultdict
from unidecode import unidecode

# Disable copy warnings
pd.options.mode.chained_assignment = None


# Determine month and year based on airbnb name
def date_from_bnb(file_name):
    # Find airbnb file
    file_full_name = ''
    for fil in os.listdir('ModelFiles/.'):
        if fnmatch.fnmatch(fil, '*.xlsx'):
            if file_name in fil:
                file_full_name = fil

    date_str = ''.join([p for p in file_full_name.split('-')[1] if p.isdigit()]).strip()
    mo = int(date_str[:2])
    yr = int(date_str[2:])

    return mo, yr


# Create function to determine reference numbers
def reference_number_read():
    try:
        with open('ref_number_values.txt', 'r') as f:
            # Load configuration file values
            invoice_line = [2]
            checks_line = [3]
            journal_line = [4]
            for position, line in enumerate(f):

                if position in invoice_line:
                    invoice_str = re.findall('\d*\.?\d+', line)
                    invoice_number = int(float(invoice_str[0]))

                elif position in checks_line:
                    checks_str = re.findall('\d*\.?\d+', line)
                    checks_number = int(float(checks_str[0]))

                elif position in journal_line:
                    journal_str = re.findall('\d*\.?\d+', line)
                    journal_number = int(float(journal_str[0]))

    except FileNotFoundError:
        # Keep preset values
        print('ref_number_values.txt not found! Choosing starting values for reference numbers.')
        print()

        invoice_number = int(input('Please input the starting reference number for Invoices: \n'))
        checks_number = int(input('Please input the starting reference number for Checks (Just the number value): \n'))
        journal_number = int(
            input('Please input the starting reference number for Journal Entries (Just the number value): \n'))

    return invoice_number, checks_number, journal_number


# Create a function to write reference numbers
def reference_number_write(invoice_no, check_no, journal_no, mo, yr):
    with open('ref_number_values.txt', 'w') as f:
        # Print date
        f.write("LIN date: " + str(mo) + "/" + str(yr) + "\n\n")
        # Load reference numbers
        f.write("Reference number (Invoice): " + f"{invoice_no}" + "\n")
        f.write("Reference number (Checks): ABB TR " + f"{check_no:05d}" + "\n")
        f.write("Reference number (Journal): PMT " + f"{journal_no:05d}" + "\n")


# Define an explanation function for missing data
def explanation_missing(df, explain):
    df.loc[0, 'Explanation'] = explain

    return df


# Use this to check if there is a nan
def isnan(value):
    try:
        return math.isnan(float(value))
    except ValueError:
        return False


# Return month name from number
def number2month(num):
    datetime_object = datetime.datetime.strptime(str(num), "%m")
    full_month_name = datetime_object.strftime("%B")

    return full_month_name


# Numbering a month
def mtn(month_full_name):
    months = {
        'jan': 1,
        'feb': 2,
        'mar': 3,
        'apr': 4,
        'may': 5,
        'jun': 6,
        'jul': 7,
        'aug': 8,
        'sep': 9,
        'oct': 10,
        'nov': 11,
        'dec': 12
    }
    a = month_full_name.strip()[:3].lower()
    ez = months[a]

    return ez


def remove_directory(directory):
    if os.path.exists(directory):
        # Close any open files in the directory
        for file in os.listdir(directory):
            file_path = os.path.join(directory, file)
            try:
                if os.path.isfile(file_path) and \
                        file_path not in map(lambda fl: fl.name, psutil.process_iter()):
                    # close only if file is not being accessed by any process
                    f = open(file_path)
                    f.close()
            except NameError:
                pass

        # Remove the directory and its contents recursively
        try:
            shutil.rmtree(directory)
            print("Directory removed successfully!")
        except OSError as e:
            print(f"Error: {directory} : {e.strerror}")
    else:
        print("Directory does not exist.")


def reformat_original_files(final, customer_df, listing_row='Listing', customer_row='ListingBNB', reformat=True):
    replace_re = "[^A-Za-z0-9_]+"
    for index, row in final.iterrows():
        listing = row[listing_row]
        if not isnan(listing):
            # Reformat the listing
            if reformat:
                listing = unidecode(listing)
                listing = listing.replace(u'\xa0', u' ')
                listing = listing.replace(replace_re, ' ')
                listing = listing.strip()
            # Find in database
            customer_info = customer_df[customer_df[customer_row] == listing]
            if not customer_info.empty:
                final['Code'].loc[index] = customer_info['Code'].iloc[0]
                final['Customer'].loc[index] = customer_info['QBO'].iloc[0]

    return final


def code_customer(app, cleaning, vrbo, bnb, check, writing_info):
    app.log('Reformatting original files to include customer codes..')

    replace_re = "[^A-Za-z0-9_]+"
    # Create a database for listing names
    cleaning = cleaning[~cleaning.ListingBNB.isnull()]
    cleaning = cleaning[~cleaning.QBO.isnull()]
    cleaning['ListingBNB'] = cleaning['ListingBNB'].apply(unidecode)
    cleaning['ListingBNB'] = cleaning['ListingBNB'].str.replace(u'\xa0', u' ')
    cleaning['ListingBNB'] = cleaning['ListingBNB'].str.replace(replace_re, ' ', regex=True)
    cleaning['ListingBNB'] = cleaning['ListingBNB'].str.strip()

    # AirBNB
    final_bnb = bnb.reindex(columns=["Code", "Customer"] + bnb.columns.tolist())
    final_bnb = reformat_original_files(final_bnb, cleaning)
    with pd.ExcelWriter(writing_info['bnb']['path'], engine='xlsxwriter') as writer:
        final_bnb.to_excel(writer, index=False, sheet_name=writing_info['bnb']['sheet'])

    # Reservations
    final_check = check.reindex(columns=["Code", "Customer"] + check.columns.tolist())
    final_check = reformat_original_files(final_check, cleaning)
    with pd.ExcelWriter(writing_info['check']['path'], engine='xlsxwriter') as writer:
        final_check.to_excel(writer, index=False, sheet_name=writing_info['check']['sheet'])

    # VRBO
    final_vrbo = vrbo.reindex(columns=["Code", "Customer"] + vrbo.columns.tolist())
    final_vrbo = reformat_original_files(final_vrbo, cleaning, 'Property ID', 'VRBO_ID', False)
    with pd.ExcelWriter(writing_info['vrbo']['path'], engine='xlsxwriter') as writer:
        final_vrbo.to_excel(writer, index=False, sheet_name=writing_info['vrbo']['sheet'])

    app.log('Reformat complete')

    return


#######################################################################################################################
# Notes
# Check check and CM dates, make sure that they are set to the first of the month.


def line_invoice_generation(app):
    app.progress_bar["value"] = 1
    app.progress_bar.update()

    ####################################################################################################################
    # Inputs

    # Which names are desired?
    des_sub_names = ['Cleaning', 'Customer']
    filenames = ['reservations', 'airbnb', 'Current', 'VRBO_']

    ####################################################################################################################
    # Determine month and year of the LIN to be run.
    month, year = date_from_bnb(filenames[1])

    # Beginning ref numbers...
    invoice_no, check_no, journal_no = reference_number_read()

    # # Date creation for invoice and due dates
    # # Find number of days in the invoice month
    # end_month = calendar.monthrange(year, month)

    # invoice = datetime.datetime(year, month, end_month[1], 0, 0)
    # invoice = invoice.strftime('%m/%d/%Y')
    # Create a string for due date
    # If December, restart month at January
    if month == 12:
        invoice = datetime.datetime(year + 1, 1, 1, 0, 0)
    else:
        invoice = datetime.datetime(year, month + 1, 1, 0, 0)
    invoice = invoice.strftime('%m/%d/%Y')

    # Create a string for due date
    # If December, restart month at January
    if month == 12:
        due = datetime.datetime(year + 1, 1, 5, 0, 0)
    else:
        due = datetime.datetime(year, month + 1, 5, 0, 0)
    due = due.strftime('%m/%d/%Y')
    month_name = number2month(month)
    # Finishing file
    finish = 'Aviad_BNB_' + month_name + '.xlsx'
    sheet_names = ['Invoices', 'Credit_Memo_Invoices', 'Credit_Memos_fields', 'Checks_fields', 'Sales_tax_fields',
                   'Sales_Receipts', 'Journal_Entries']

    # Item descriptions that are currently desired
    item = ['CLEANING FEE', 'HOSPITALITY TAX', 'MANAGEMENT FEE']

    # Create empty dataframes for each desired excel sheet
    bnb = pd.DataFrame()
    cleaning = pd.DataFrame()
    customer_info = pd.DataFrame()
    check = pd.DataFrame()
    vrbo = pd.DataFrame()

    # Create empty dataframes for each set of data to be written to output file
    entry_NCM = pd.DataFrame()
    entry_CM = pd.DataFrame()
    sales_entry = pd.DataFrame()
    credit_memo = pd.DataFrame()
    sales_receipts = pd.DataFrame()
    journal_entries = pd.DataFrame()
    checks = pd.DataFrame()
    tax_issues = pd.DataFrame()
    man_issues = pd.DataFrame()
    prev_cleaning = pd.DataFrame()

    ####################################################################################################################
    # Data Import
    app.progress_bar["value"] = 2
    app.progress_bar.update()

    # Create necessary directories if not already done so.
    directory_month = month_name + ' Report'
    cwd = os.getcwd()
    path = cwd + '\\' + f'{year} Reports' + '\\' + directory_month
    remove_directory(path)
    filepath = cwd + '\\ModelFiles'

    # Check whether the specified path exists or not
    is_exist = os.path.exists(path)

    if not is_exist:
        # Create a new directory because it does not exist
        os.makedirs(path)
        app.log('The new directory is created!')

    # Load the necessary files
    app.log("Loading Directory Files...")
    app.progress_bar["value"] = 3
    app.progress_bar.update()
    num_files = len(os.listdir(filepath))
    progress_vals = np.linspace(3, 15, num=num_files)
    file_idx = 0
    reformat_info = {
        'bnb': {'path': '', 'sheet': ''},
        'check': {'path': '', 'sheet': ''},
        'vrbo': {'path': '', 'sheet': ''}
    }
    for fil in os.listdir(filepath):
        app.progress_bar["value"] = int(progress_vals[file_idx])
        app.progress_bar.update()
        file_idx += 1

        file = filepath + '\\' + fil
        for fname in range(len(filenames)):
            if fil.lower().startswith(filenames[fname].lower()):
                if fname == 0:
                    check = pd.read_excel(file)
                    reformat_info['check']['path'] = path + '\\' + fil
                    reformat_info['check']['sheet'] = fil
                elif fname == 1:
                    bnb = pd.read_excel(file)
                    reformat_info['bnb']['path'] = path + '\\' + fil
                    reformat_info['bnb']['sheet'] = fil
                elif fname == 2:
                    xl = pd.ExcelFile(file)

                    for i in range(len(des_sub_names)):
                        # Define sheet name
                        sheet = [string for string in xl.sheet_names if des_sub_names[i] in string]

                        if i == 0:
                            # Determine correct sheet name and put in string format.
                            name = [name for name in xl.sheet_names if name in sheet]
                            name_str = ''.join(name)

                            # Create DataFrame
                            cleaning = pd.read_excel(file, sheet_name=name_str)
                        elif i == 1:
                            # Determine correct sheet name and put in string format.
                            name = [name for name in xl.sheet_names if name in sheet]
                            name_str = ''.join(name)

                            # Create DataFrame
                            customer_info = pd.read_excel(file, sheet_name=name_str)

                            # Copy Running Customer list to Report Folder
                            copyfile(filepath + '\\' + fil, path + '\\' + fil)
                elif fname == 3:
                    try:
                        vrbo = pd.read_excel(file)
                    except ValueError:
                        vrbo = pd.read_csv(file)
                    reformat_info['vrbo']['path'] = path + '\\' + fil
                    reformat_info['vrbo']['sheet'] = fil

    # For VRBO, AirBNB, and Reservations append Code and Customer in first columns
    code_customer(app, cleaning, vrbo, bnb, check, reformat_info)

    # Remove any rows including "Resolution" because they are pet fees and should not be included.
    bnb = bnb[~bnb['Type'].isin(['Resolution Payout'])]
    bnb = bnb[~bnb['Type'].isin(['Resolution Adjustment'])]

    ####################################################################################################################
    # Create data frames of information'
    app.log("Optimizing Data...")
    app.progress_bar["value"] = 16
    app.progress_bar.update()

    # Find the unit you want and determine all important information
    # Begin with the columns of the information desired
    replace_re = "[^A-Za-z0-9_ -:&]+"
    bnb_col = pd.DataFrame(bnb, columns=['Listing', 'Amount', 'Type', 'Confirmation Code', 'Nights'])
    bnb_col.loc[bnb_col["Listing"].isnull(), 'Listing'] = 'NULL'
    bnb_col['Listing'] = bnb_col['Listing'].apply(unidecode)
    bnb_col['Listing'] = bnb_col['Listing'].str.replace(u'\xa0', u' ')
    bnb_col = bnb_col[~bnb_col.Listing.isnull()]
    bnb_col['Listing'] = bnb_col['Listing'].str.replace(replace_re, ' ', regex=True)
    bnb_col['Listing'] = bnb_col['Listing'].str.strip()
    # both_col = pd.DataFrame(both, columns=['ListingBNB', 'QBO', 'Cleaning', 'Tax_Location'])
    cleaning_col = pd.DataFrame(cleaning,
                                columns=['ListingBNB', 'QBO', 'Cleaning', 'Tax_Location', 'Pest', 'Landscape',
                                         'Internet/Cable', 'Bus_Lic', 'VRBO_ID', 'Code', 'Output'])
    cleaning_col.loc[cleaning_col["ListingBNB"].isnull(), 'ListingBNB'] = 'NULL'
    cleaning_col['ListingBNB'] = cleaning_col['ListingBNB'].apply(unidecode)
    cleaning_col['ListingBNB'] = cleaning_col['ListingBNB'].str.replace(u'\xa0', u' ')
    cleaning_col = cleaning_col[~cleaning_col.ListingBNB.isnull()]
    cleaning_col = cleaning_col[~cleaning_col.QBO.isnull()]
    cleaning_col['ListingBNB'] = cleaning_col['ListingBNB'].str.replace(replace_re, ' ', regex=True)
    cleaning_col['ListingBNB'] = cleaning_col['ListingBNB'].str.strip()
    cleaning_col['QBO'] = cleaning_col['QBO'].str.replace(replace_re, ' ', regex=True)
    cleaning_col['QBO'] = cleaning_col['QBO'].str.strip()
    customer_col = pd.DataFrame(customer_info,
                                columns=['Customer-QBO', 'Expense_Flat', 'Credit', 'Clean', 'Hosp', 'Management',
                                         'Magpercent'])
    customer_col['Customer-QBO'] = customer_col['Customer-QBO'].str.replace(u'\xa0', u' ')
    customer_col = customer_col[~customer_col['Customer-QBO'].isnull()]
    customer_col['Customer-QBO'] = customer_col['Customer-QBO'].str.replace(replace_re, ' ', regex=True)
    customer_col['Customer-QBO'] = customer_col['Customer-QBO'].str.strip()
    check_col = pd.DataFrame(check, columns=['Listing'])
    check_col.loc[check_col["Listing"].isnull(), 'Listing'] = 'NULL'
    check_col['Listing'] = check_col['Listing'].apply(unidecode)
    check_col['Listing'] = check_col['Listing'].str.replace(replace_re, ' ', regex=True)
    check_col['Listing'] = check_col['Listing'].str.strip()
    vrbo_col = pd.DataFrame(vrbo, columns=['Property ID', 'Reservation ID', 'Payout', 'Nights', 'Check-out'])

    # VRBO in different months
    vrbo_wrong_month = vrbo[vrbo['Check-out'].dt.month != month]
    vrbo_save_path = filepath + '\\VRBO Date Organizer.xlsx'

    # Add any old vrbo listings to this month's report
    file_exists = exists(vrbo_save_path)
    vrbo_new = pd.DataFrame()
    keep_locations = []
    if file_exists:
        vrbo_new = pd.read_excel(vrbo_save_path)
        # Iterate over new vrbo that is not in current month. If the reservation is not in the current list, add it
        for idx, res_ID in enumerate(vrbo_wrong_month['Reservation ID']):
            if (vrbo_new['Reservation ID'] != res_ID).all():
                vrbo_new = pd.concat([vrbo_new, vrbo_wrong_month.iloc[[idx]]], ignore_index=True, axis=0)
                # add month and year
                vrbo_new.iloc[-1, 0] = year
                vrbo_new.iloc[-1, 1] = number2month(vrbo_wrong_month['Check-out'].dt.month.iloc[idx])

        # Create locations of vrbo listings that do not need to be deleted
        for idx, v_month in enumerate(vrbo_new['Month']):
            current_month_number = mtn(month_name)
            if mtn(v_month) >= current_month_number and vrbo_new.iloc[idx, 0] >= year:
                keep_locations.append(idx)
        vrbo_new = vrbo_new.iloc[keep_locations]
        vrbo_new.loc[vrbo_new['Payout'] != 0, 'Payout'] = 0

        # Add any reservations checking out this month to current vrbo listings
        if not vrbo_new.empty:
            try:
                vrbo_old = vrbo_new
                vrbo_old = vrbo_old.loc[vrbo_old['Month'] == month_name]
                vrbo_old = vrbo_old[['Property ID', 'Reservation ID', 'Payout', 'Nights', 'Check-out']]
                # Add listings to vrbo
                vrbo_col = pd.concat([vrbo_col, vrbo_old], ignore_index=True, axis=0)
            except ValueError:
                app.log('Old VRBO listings not found. Moving on..')

    # Write back to excel
    os.remove(vrbo_save_path)
    vrbo_new.to_excel(vrbo_save_path, sheet_name='VRBO Check Outs', index=False)

    # Copy Running Customer list to Report Folder
    copyfile(vrbo_save_path, path + '\\' + 'VRBO Date Organizer.xlsx')
    ###################################################################################################################
    # Missing information and general housekeeping
    app.log("Finding Missing Information...")

    # Create a list of Listings that do not have a customer attached to them.
    diff = list(set(bnb_col['Listing']) - set(cleaning_col['ListingBNB']))

    # If there are values in diff, create a dataframe of information for this case
    listing_diff = pd.DataFrame()
    for i in diff:
        diff_new = bnb[bnb['Listing'] == i]
        listing_diff = pd.concat([listing_diff, diff_new], ignore_index=True)

    # Found issue with spacing. Fix by making sure there is only 1 space in all names
    bnb_col['Listing'] = bnb_col['Listing'].str.replace('  ', ' ')
    cleaning_col['ListingBNB'] = cleaning_col['ListingBNB'].str.replace('  ', ' ')
    check_col['Listing'] = check_col['Listing'].str.replace('  ', ' ')

    # # Determine if there are discrepancies in the data based on customer names and listings and report them
    # # Differences between both and qbo customer names could simply mean the names are slightly different and need to
    # # be fixed.
    # customer_diff_both = list(set(customer_col['Customer-QBO']) - set(cleaning_col['QBO']))
    # for i in customer_diff_both:
    #     print('The customer ' + i + ' is not listed in Cleaning Fee Report')

    customer_diff_qbo = list(set(cleaning_col['QBO']) - set(customer_col['Customer-QBO']))
    both_empty = cleaning_col['ListingBNB'].isnull()
    both_index = cleaning_col[both_empty]
    if not both_index.empty:
        customer_diff_qbo = pd.concat([customer_diff_qbo, both_index['QBO'].iloc[0]])
    customer_diff = pd.DataFrame()
    if bool(customer_diff_qbo):
        print('The customer ' + customer_diff_qbo[
            0] + ' is not listed in Customer Report. This needs to be fixed if you would like the code to operate as '
                 'it should.')

        for i in customer_diff_qbo:
            customer_diff_new = cleaning[cleaning['QBO'] == i]
            customer_diff = pd.concat([customer_diff, customer_diff_new], ignore_index=True)

    # Determine differences in VRBO
    vrbo_diff_id = list(set(vrbo_col['Property ID']) - set(cleaning_col['VRBO_ID']))
    vrbo_diff = pd.DataFrame()
    for v in vrbo_diff_id:
        vrbo_diff_new = vrbo[vrbo['Property ID'] == v]
        vrbo_diff = pd.concat([vrbo_diff, vrbo_diff_new], ignore_index=True)

    ###################################################################################################################
    # Separate Aviad's listings
    app.log("Separating Aviad's Listings...")
    aviad_excel = defaultdict(dict)
    for index, row in cleaning_col.iterrows():
        try:
            # For any case where the cell is a string
            if len(row['Output']) > 2 and bool(re.match(r'(\w+)', row['Output'])):
                if not bool(aviad_excel[row['Output']]):
                    aviad_excel[row['Output']] = {}
                    aviad_excel[row['Output']]['listings2find'] = []
                    aviad_excel[row['Output']]['vrbo_ids'] = []

                aviad_excel[row['Output']]['listings2find'].append(row['ListingBNB'])
                aviad_excel[row['Output']]['vrbo_ids'].append(row['VRBO_ID'])
        except TypeError:
            pass  # The cell is empty and a nan is returned

    # Filter and export
    for name in aviad_excel:
        filtered_bnb = bnb[bnb['Listing'].isin(aviad_excel[name]['listings2find'])]
        filtered_vrbo = vrbo[vrbo['Property ID'].isin(aviad_excel[name]['vrbo_ids'])]
        with pd.ExcelWriter(path + '\\' + name + ' Reservations ' + str(month) + '_' + str(year) + '.xlsx',
                            engine='xlsxwriter') as f_writer:
            filtered_bnb.to_excel(f_writer, index=False, sheet_name='AirBNB')
            filtered_vrbo.to_excel(f_writer, index=False, sheet_name='VRBO')

    ###################################################################################################################
    # Create total unit containing all necessary data

    # Create a lookup table for each customer including listings, income, cleaning fees, etc.
    unit = pd.DataFrame()
    num_customers = len(customer_col['Customer-QBO'])
    progress_vals = np.linspace(16, 25, num=num_customers)
    for customers in range(len(customer_col['Customer-QBO'])):
        app.progress_bar["value"] = int(progress_vals[customers])
        app.progress_bar.update()
        # Find customer
        customer = customer_col['Customer-QBO'].iloc[customers]
        # Find all instances of this customer
        customer_listing = cleaning_col[cleaning_col['QBO'] == customer]
        for listings in range(len(customer_listing['QBO'])):
            vrbo_90_day_amount = 0
            bnb_90_day_amount = 0

            # Find all listings under this customer
            listing = bnb_col[bnb_col['Listing'] == customer_listing['ListingBNB'].iloc[listings]]
            if listing.empty:
                # Add the income and checkouts
                listing.at[0, 'Listing'] = customer_listing['ListingBNB'].iloc[listings]
                listing.at[0, 'Amount'] = 0
                listing.at[0, 'Type'] = ''
                listing.at[0, 'Confirmation Code'] = ''
                listing.at[0, 'Nights'] = 0

                income = 0
                reservation = check_col[check_col['Listing'] == customer_listing['ListingBNB'].iloc[listings]]
                number_reservations = len(reservation['Listing'])
            else:
                # Determine amount to not include in taxes due to 90 day term rentals
                for idx, num_nights in enumerate(listing['Nights']):
                    if num_nights > 89:
                        bnb_90_day_amount += listing['Amount'].iloc[idx]

                # Add the income and checkouts
                listings_minus_passthroughs = listing[~listing['Type'].str.contains('Pass', case=False, regex=False)]
                income = listings_minus_passthroughs['Amount'].sum()
                reservation = check_col[check_col['Listing'] == listing['Listing'].iloc[0]]
                number_reservations = len(reservation['Listing'])

            # Determine if VRBO, and find payout
            vrbo_nights = 0
            vrbo_id = 'none'
            vrbo_payout = 0
            if not isnan(customer_listing['VRBO_ID'].iloc[listings]):
                vrbo_id = customer_listing['VRBO_ID'].iloc[listings]
                vrbo_payouts = vrbo_col[vrbo_col['Property ID'] == vrbo_id]
                if vrbo_payouts.empty:
                    vrbo_payout = 0
                else:
                    vrbo_payout = vrbo_payouts['Payout'].sum()
                    vrbo_nights = len(set(vrbo_payouts['Reservation ID']))

                    # Determine amount to not include in taxes due to 90 day term rentals
                    for idx, num_nights in enumerate(vrbo_payouts['Nights']):
                        if num_nights > 89:
                            vrbo_90_day_amount += vrbo_payouts['Payout'].iloc[idx]

                # Do not count cleaning fees for reservations that end in the incorrect month
                incorrect_month = (vrbo_payouts['Check-out'].dt.month != month).value_counts()
                try:
                    incorrect_month_count = incorrect_month.loc[True]
                except KeyError:
                    incorrect_month_count = 0
                vrbo_nights = vrbo_nights - incorrect_month_count

                # Create dataframe
            unit_new = pd.DataFrame(
                [customer,
                 listing['Listing'].iloc[0],
                 round(income, 2),
                 customer_listing['Cleaning'].iloc[listings],
                 number_reservations,
                 customer_listing['Tax_Location'].iloc[listings],
                 customer_listing['Pest'].iloc[listings],
                 customer_listing['Landscape'].iloc[listings],
                 customer_listing['Internet/Cable'].iloc[listings],
                 customer_listing['Bus_Lic'].iloc[listings],
                 customer_col['Expense_Flat'].iloc[customers],
                 customer_col['Credit'].iloc[customers],
                 customer_col['Clean'].iloc[customers],
                 customer_col['Hosp'].iloc[customers],
                 customer_col['Management'].iloc[customers],
                 customer_col['Magpercent'].iloc[customers],
                 vrbo_id, vrbo_payout, vrbo_nights,
                 customer_listing['Code'].iloc[listings],
                 bnb_90_day_amount, vrbo_90_day_amount])
            unit = pd.concat([unit, unit_new.transpose()], ignore_index=True)

    # Create field names for entry
    unit.columns = ['Customer', 'Listing', 'Income', 'CleaningFee', 'Checkouts', 'TaxLocation', 'Pest',
                    'Landscape', 'Internet/Cable', 'Bus_Lic', 'Expense', 'CreditMemo', 'Clean', 'Hosp', 'Management',
                    'Magpercent',
                    'VRBO_ID',
                    'VRBO_PAYOUT', 'VRBO_Nights', 'Code', 'BNB_90_Day', 'VRBO_90_Day']

    ###################################################################################################################
    # Data Extraction
    app.log("=== Beginning Invoicing ===", True)

    # Sort unit
    unit = unit.sort_values(by=['Customer'])
    unit_names = unit['Customer'].unique()

    # If the cleaning fee is a nan, change it to 0
    unit['CleaningFee'] = unit['CleaningFee'].apply(pd.to_numeric, errors='coerce')
    unit['CleaningFee'] = unit['CleaningFee'].fillna(0)

    unit.loc[unit["CreditMemo"].isnull(), 'CreditMemo'] = 'NULL'
    unit_grouped = unit.groupby('CreditMemo')

    num_invoices = len(unit_names)
    progress_vals = np.linspace(25, 95, num=num_invoices)

    for ug in range(len(unit_grouped.groups.keys())):
        entry = pd.DataFrame()
        if ug == 0:
            unit = unit_grouped.get_group('CM')
        elif ug == 1:
            unit = unit_grouped.get_group('NULL')
        else:
            app.log('There was an error separating by CM!!! Have Steven look into it please!')

        for n in range(len(unit_names)):

            app.progress_bar["value"] = int(progress_vals[n])
            app.progress_bar.update()

            # InvoiceNo
            invoice_no += 1

            # Separate by name
            unit_loop = unit[unit['Customer'] == unit_names[n]]

            # Customer name
            try:
                customer_name = unit_loop['Customer'].iloc[0]
            except IndexError:
                # Because of Credit Memo split, this may be in the other section. Just move on
                continue

            # Determine the total income
            income_total_init = unit_loop['Income'].sum()

            # If VRBO, add to income but keep separate for tax information
            if not all(unit_loop['VRBO_ID'] == 'none'):
                income_total = income_total_init + unit_loop['VRBO_PAYOUT'].sum()
            else:
                income_total = income_total_init

            # Check if un-invoiced
            un_invoiced = False
            # If all entries say omit, this is un-invoiced
            if unit_loop['Clean'].iloc[0] == 'omit' and unit_loop['Hosp'].iloc[0] == 'omit' and \
                    unit_loop['Management'].iloc[
                        0] == 'omit':
                un_invoiced = True

            # Initialization
            total_cleaning_fee = 0
            total_amount = 0
            # management_fee = 0
            pest_fee = 0
            bus_lic = 0
            landscaping_fee = 0
            cable_fee = 0
            expense = 0
            clean_count = 0
            cleaning_written = False
            # hospitality_written = False
            tax_adjustment_90_day = 0

            # Separate based on cleaning fee and tax location
            unique_cleaning = unit_loop['CleaningFee'].unique()
            unique_tax = unit_loop['TaxLocation'].unique()

            for tax in unique_tax:
                # Get Rid Of Later
                # if code

                if not isnan(tax):
                    unit_tax = unit_loop[unit_loop['TaxLocation'] == tax]
                else:
                    unit_tax = unit_loop[unit_loop['TaxLocation'].isna()]

                for clean in unique_cleaning:
                    unit_repeat = unit_tax[unit_tax['CleaningFee'] == clean]
                    if unit_repeat.empty:
                        continue

                    # Initialize necessary values
                    cleaning_fee = 0
                    num_cleaning_fee = 0
                    amount_less_clean = 0
                    hospitality_tax = 0
                    tax_adjustment_90_day = 0

                    # Cleaning fee
                    if not isnan(unit_repeat['CleaningFee'].iloc[0]) or \
                            not unit_repeat['Clean'].str.contains('omit', case=False).any():
                        cleaning_fee = unit_repeat['CleaningFee'].iloc[0]
                        num_cleaning_fee = unit_repeat['Checkouts'].sum()
                        # Add a cleaning fee for VRBO reports
                        num_cleaning_fee = num_cleaning_fee + unit_repeat['VRBO_Nights'].sum()
                        total_cleaning_fee = total_cleaning_fee + cleaning_fee * num_cleaning_fee

                    # Hospitality tax
                    if not unit_repeat['Clean'].str.contains('omit', regex=False, case=False).any():
                        tax_adjustment_90_day = unit_repeat['BNB_90_Day'].sum() + unit_repeat['VRBO_90_Day'].sum()
                        amount_less_clean = unit_repeat[
                                                'Income'].sum() - cleaning_fee * num_cleaning_fee - tax_adjustment_90_day
                        total_amount += amount_less_clean
                        hospitality_tax = 0.03 * amount_less_clean

                    # If the customer has nothing to invoice, skip it
                    if total_cleaning_fee == 0 and unit_repeat['Income'].sum() == 0:
                        continue

                    # Log sales taxes
                    if not hospitality_tax == 0 and not isnan(tax):
                        sales_entry_new = pd.DataFrame(
                            [tax, round(amount_less_clean, 2), round(hospitality_tax / 2, 2),
                             round(hospitality_tax / 2, 2),
                             round(hospitality_tax, 2)])
                        sales_entry = pd.concat([sales_entry, sales_entry_new.transpose()], ignore_index=True)
                    elif isnan(tax):
                        tax_issues = pd.concat([tax_issues, unit_repeat], ignore_index=True)

                    # Log Cleaning Fee
                    if not un_invoiced:
                        # Was there a cleaning fee already reported for this customer? If so add this to what is listed
                        # already.
                        if not entry.empty:
                            prev_cleaning = entry[entry[1].str.contains(customer_name, regex=False)]
                            prev_cleaning = prev_cleaning[prev_cleaning[5].str.contains('CLEANING', regex=False)]
                            prev_cleaning = prev_cleaning[prev_cleaning[8] == clean]

                        if prev_cleaning.empty:
                            if not unit_repeat['Clean'].str.contains('del', regex=False, case=False).any():
                                cleaning_written = True
                                clean_count += 1
                                entry_new = pd.DataFrame(
                                    [invoice_no, unit_repeat['Customer'].iloc[0], invoice, due, '', item[0], item[0],
                                     num_cleaning_fee,
                                     round(cleaning_fee, 2), round(cleaning_fee * num_cleaning_fee, 2), tax, invoice])
                                entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)
                        else:
                            cleaning_written = True
                            new_cleaning_number = entry.loc[prev_cleaning.index, 7].iloc[0] + num_cleaning_fee
                            new_cleaning_amount = new_cleaning_number * clean
                            entry.loc[prev_cleaning.index, 7] = new_cleaning_number
                            entry.loc[prev_cleaning.index, 9] = new_cleaning_amount

                # If any fee is below 0, then make them 0.
                total_amount = max(0, total_amount)

                # If the customer has nothing to invoice, skip it
                if total_cleaning_fee == 0 and total_amount == 0:
                    continue

            # If the customer has nothing to invoice, skip it
            if total_cleaning_fee == 0 and total_amount == 0 and tax_adjustment_90_day == 0 \
                    and unit_loop['VRBO_PAYOUT'].sum() == 0:
                continue

            if not isnan(unit_loop['Expense'].iloc[0]):
                expense = unit_loop['Expense'].iloc[0]
            # Management fee
            total_management = income_total - total_cleaning_fee - expense
            man_rate = unit_loop['Magpercent'].iloc[0]
            if isnan(man_rate):
                if unit_loop['Management'].str.contains('omit', regex=False, case=False).any():
                    man_rate = 0.00
                else:
                    man_issues = pd.concat([man_issues, unit_loop], ignore_index=True)
            management_fee = 0.01 * man_rate * total_management
            if management_fee < 0.00:
                management_fee = 0.00

            # If any fee is below 0, then make them 0.
            total_management = max(0, total_management)

            # Create the memo
            # Change formatting of memo values
            income_memo = '${0:,.2f}'.format(round(income_total, 2))
            expense_memo = '${0:,.2f}'.format(round(expense, 2))
            cleaning_memo = '${0:,.2f}'.format(round(total_cleaning_fee, 2))
            total_memo = '${0:,.2f}'.format(round(total_management, 2))
            manage_memo = '${0:,.2f}'.format(round(management_fee, 2))

            # If cleaning fee or expenses are 0, do not include in memo
            memo = 'INCOME ' + income_memo
            if total_cleaning_fee != 0:
                memo += ' - CLEANING ' + cleaning_memo
            if expense != 0:
                memo += ' - EXPENSES ' + expense_memo
            memo += ' = ' + total_memo + ' | MANAGEMENT FEE ' + str(man_rate) + '% --> ' + manage_memo

            # Place the memo in the cleaning fee and hospitality tax sections
            if cleaning_written:
                entry.iloc[-1, 4] = memo
                for cl in (n + 2 for n in range(clean_count - 1)):
                    entry.iloc[-cl, 4] = memo

            # Log Management Fee
            tax_location = entry.iloc[-1, 10]
            if not unit_loop['Management'].iloc[0] == 'omit':
                if not unit_loop['Management'].str.contains('del', regex=False, case=False).any():
                    entry_new = pd.DataFrame(
                        [invoice_no, customer_name, invoice, due, memo, item[2], item[2], 0.01 * man_rate,
                         round(total_management, 2), round(management_fee, 2), tax_location, invoice])
                    entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Log Pest Control Fee
            if not unit_loop['Pest'].isnull().values.all():
                pest_fee = unit_loop['Pest'].sum()
                entry_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice, due, memo, 'SERVICES', 'PEST CONTROL', 1,
                     round(pest_fee, 2), round(pest_fee, 2), tax_location, invoice])
                entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Log Landscaping Fee
            if not unit_loop['Landscape'].isnull().values.all():
                landscaping_fee = unit_loop['Landscape'].sum()
                entry_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice, due, memo, 'SERVICES', 'LANDSCAPING', 1,
                     round(landscaping_fee, 2), round(landscaping_fee, 2), tax_location, invoice])
                entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Log Internet/Cable Fee
            if not unit_loop['Internet/Cable'].isnull().values.all():
                cable_fee = unit_loop['Internet/Cable'].sum()
                entry_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice, due, memo, 'SERVICES', 'INTERNET_CABLE', 1,
                     round(cable_fee, 2), round(cable_fee, 2), tax_location, invoice])
                entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Log Pest Control Fee
            if not unit_loop['Bus_Lic'].isnull().values.all():
                bus_lic = unit_loop['Bus_Lic'].sum()
                entry_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice, due, memo, 'SERVICES', 'BUSINESS LICENSE', 1,
                     round(bus_lic, 2), round(bus_lic, 2), tax_location, invoice])
                entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Credit Memo/Checks
            if (unit_loop['CreditMemo'] == 'CM').any():
                # Create item description for Credit Memo
                item_cm = 'Deposits in Trust'
                item_cm_desc = 'Funds Collected on Behalf of Client'

                # Log Credit Memo
                credit_memo_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice, invoice, item_cm, item_cm_desc,
                     1, round(income_total, 2), round(income_total, 2), tax_location])
                credit_memo = pd.concat([credit_memo, credit_memo_new.transpose()], ignore_index=True)

                # Sales Receipts
                item_sr = 'TRUST Clearing Account'
                bnb_sales = False
                vrbo_sales = False
                bnb_income = unit_loop['Income'].sum()
                vrbo_income = unit_loop['VRBO_PAYOUT'].sum()
                total_sale_amt = round(bnb_income + vrbo_income, 2)
                if not bnb_income == 0.0:
                    bnb_sales = True
                if not vrbo_income == 0.0:
                    vrbo_sales = True
                if bnb_sales and not vrbo_sales:
                    pmnt_meth = 'AirBNB'
                elif not bnb_sales and vrbo_sales:
                    pmnt_meth = 'VRBO'
                elif bnb_sales and vrbo_sales:
                    pmnt_meth = 'BOTH'
                else:
                    pmnt_meth = ''

                bnb_income = '${0:,.2f}'.format(round(bnb_income, 2))
                vrbo_income = '${0:,.2f}'.format(round(vrbo_income, 2))
                sales_msg = f'Receipt income in trust from - AirBNB Amount: {bnb_income} | VRBO Amount: {vrbo_income}'
                line_itm = 'Rents in trust - Liability'

                # Log Sales Receipts
                sr_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice, tax_location, item_sr, pmnt_meth,
                     sales_msg, 'N', 'N', invoice, line_itm, sales_msg, 1, total_sale_amt, total_sale_amt, 'NON'])
                sales_receipts = pd.concat([sales_receipts, sr_new.transpose()], ignore_index=True)

                # Journal Entries
                ref_num = 'PMT ' + f'{journal_no:05d}'
                customer = unit_loop['Customer'].iloc[0]
                prvt_note = f'Payment {customer} to #{invoice_no}'
                debit_acct = 'Accounts Receivable (A/R)'
                credit_acct = 'ABB Trust #5241'

                # Log Journal Entries
                # Debit
                je_new = pd.DataFrame(
                    [ref_num, invoice, prvt_note, 'False', debit_acct, income_total, debit_acct, tax_location,
                     customer])
                journal_entries = pd.concat([journal_entries, je_new.transpose()], ignore_index=True)
                # Credit
                je_new = pd.DataFrame(
                    [ref_num, invoice, prvt_note, 'False', credit_acct, -income_total, credit_acct, tax_location,
                     customer])
                journal_entries = pd.concat([journal_entries, je_new.transpose()], ignore_index=True)
                journal_no += 1

                # Checks
                # Reference number
                ref_num = 'ABB TR ' + f'{check_no:05d}'
                # Bank Account
                bank = 'ABB Trust #5241'

                # Expense descriptions
                expense_desc_object = datetime.datetime.strptime(str(month), '%m')
                expense_desc_month = expense_desc_object.strftime('%B')
                expense_desc = expense_desc_month + ' Earning'
                expense_acc = 'Accounts Receivable (A/R)'

                # Create an expense amount by taking the total income and subtracting the total cleaning fees, the
                # hospitality tax and the management fee.
                total_invoice = round(
                    total_cleaning_fee + management_fee + pest_fee + bus_lic + landscaping_fee + cable_fee,
                    2)
                expense_amt = income_total - total_invoice

                # Memo
                income_check_memo = '${0:,.2f}'.format(round(income_total, 2))
                total_invoice_memo = '${0:,.2f}'.format(round(total_invoice, 2))
                expense_check_memo = '${0:,.2f}'.format(round(expense_amt, 2))
                private_note = 'INCOME ' + income_check_memo + ' - CREDIT MEMO APPLIED TO INVOICE ' + total_invoice_memo + \
                               ' = MONTHLY EARNINGS ' + expense_check_memo

                checks_new = pd.DataFrame(
                    [ref_num, bank, invoice, customer_name, expense_amt, private_note, expense_desc,
                     expense_acc])
                checks = pd.concat([checks, checks_new.transpose()], ignore_index=True)
                check_no += 1

        if ug == 0:
            entry_CM = entry
        elif ug == 1:
            entry_NCM = entry
        else:
            app.log('There was an error separating by CM!!! Have Steven look into it please!')

    # Save the final invoice number
    reference_number_write(invoice_no - 1, check_no, journal_no, month, year)

    # Create field names for entry and sales_entry
    entry_CM.columns = ['RefNumber', 'Customer', 'TxnDate', 'DueDate', 'Msg', 'LineItem', 'LineDesc', 'LineQty',
                        'LineUnitPrice', 'LineAmount', 'Location', 'LineServiceDate']
    entry_NCM.columns = ['RefNumber', 'Customer', 'TxnDate', 'DueDate', 'Msg', 'LineItem', 'LineDesc', 'LineQty',
                         'LineUnitPrice', 'LineAmount', 'Location', 'LineServiceDate']
    sales_entry.columns = ['Tax_Location', 'Income', 'Municipality', 'County', 'Hospitality_Tax']

    # Create field names for credit_memo and checks
    credit_memo.columns = ['RefNumber', 'Customer', 'TxnDate', 'LineServiceDate', 'LineItem', 'LineDesc',
                           'LineQty', 'LineUnitPrice', 'LineAmount', 'Location']
    checks.columns = ['RefNumber', 'BankAccount', 'TxnDate', 'Vendor', 'ExpenseAmount', 'PrivateNote', 'ExpenseDesc',
                      'ExpenseAccount']
    sales_receipts.columns = ['RefNumber', 'Customer', 'TxnDate', 'Location', 'BankAccount', 'PaymentMethod', 'Msg',
                              'ToBePrinted', 'ToBeEmailed', 'LineServiceDate', 'LineItem', 'LineDesc', 'LineQty',
                              'LineUnitPrice', 'LineAmount', 'LineTaxable']
    journal_entries.columns = ['RefNumber', 'TxnDate', 'PrivateNote', 'IsAdjustment', 'Account', 'LineAmount',
                               'LineDesc', 'Location', 'Customer']

    # Change data type and apply formatting to specific columns
    entry_CM['LineUnitPrice'] = entry_CM['LineUnitPrice'].astype(float)  # ItemRate
    entry_CM['LineAmount'] = entry_CM['LineAmount'].astype(float)  # ItemAmount
    entry_NCM['LineUnitPrice'] = entry_NCM['LineUnitPrice'].astype(float)  # ItemRate
    entry_NCM['LineAmount'] = entry_NCM['LineAmount'].astype(float)  # ItemAmount
    credit_memo['LineUnitPrice'] = credit_memo['LineUnitPrice'].astype(float)  # ItemRate
    credit_memo['LineAmount'] = credit_memo['LineAmount'].astype(float)  # ItemAmount
    checks['ExpenseAmount'] = checks['ExpenseAmount'].astype(float)  # Expense Amount

    # For the sales sheet, create a unique list of Tax Locations and sum all corresponding elements
    unique_sales = sales_entry['Tax_Location'].unique()

    fin_sales = pd.DataFrame()
    num_updates = len(unique_sales)
    progress_vals = np.linspace(95, 100, num=num_updates)
    update_idx = 0
    for i in unique_sales:
        app.progress_bar["value"] = int(progress_vals[update_idx])
        app.progress_bar.update()
        update_idx += 1

        sales_idx = sales_entry[sales_entry['Tax_Location'] == i]

        # Sum the values
        s_inc = sales_idx['Income'].sum()
        s_mun = sales_idx['Municipality'].sum()
        s_co = sales_idx['County'].sum()
        s_hosp = sales_idx['Hospitality_Tax'].sum()
        t_loc = sales_idx['Tax_Location'].iloc[0]

        # Create the final sales data array from these values
        new_s = pd.DataFrame([t_loc, s_inc, s_mun, s_co, s_hosp])
        fin_sales = pd.concat([fin_sales, new_s.transpose()], ignore_index=True)

    fin_sales.columns = sales_entry.columns

    # When writing to excel, format the data columns with accounting format for easier viewing.
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(path + '\\' + finish, engine='xlsxwriter')

    with pd.ExcelWriter(path + '\\' + finish, engine='xlsxwriter') as writer:
        # Convert the dataframe to an XlsxWriter Excel object.
        for i in sheet_names:

            # Define a function to automatically fit columns
            def get_col_widths(df, sn):

                xl_sht_name = str(sn)

                for col in df:
                    column_wid = max(df[col].astype(str).map(len).max(), len(col))
                    col_index = df.columns.get_loc(col)
                    writer.sheets[xl_sht_name].set_column(col_index, col_index, math.ceil(column_wid * 1.25))

            if i == 'Invoices':
                entry_NCM.to_excel(writer, index=False, sheet_name=i)
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(entry_NCM, i)

            elif i == 'Credit_Memo_Invoices':
                entry_CM.to_excel(writer, index=False, sheet_name=i)
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(entry_CM, i)

            elif i == 'Credit_Memos_fields':
                credit_memo.to_excel(writer, index=False, sheet_name=i)
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(credit_memo, i)

            elif i == 'Checks_fields':
                checks.to_excel(writer, index=False, sheet_name=i)
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(checks, i)

            elif i == 'Sales_tax_fields':
                fin_sales.to_excel(writer, index=False, sheet_name=i)

                # Get the xlsxwriter workbook and worksheet objects.
                workbook = writer.book
                worksheet = writer.sheets[i]

                # Add some cell formats.
                format_col = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})

                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Format Accounting cells
                worksheet.set_column(1, 4, 18, format_col)
                worksheet.set_column(0, 0, 12)

            elif i == 'Sales_Receipts':
                sales_receipts.to_excel(writer, index=False, sheet_name=i)
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(sales_receipts, i)

            elif i == 'Journal_Entries':
                journal_entries.to_excel(writer, index=False, sheet_name=i)
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(journal_entries, i)

            # If there needs to be a sheet containing missing information, place here:
            if not listing_diff.empty:
                # Provide an explanation
                list_exp = 'Listing was found in bnb but it was not found in Cleaning Fee Report'
                listing_diff = explanation_missing(listing_diff, list_exp)

                listing_diff.to_excel(writer, index=False, sheet_name='Missing Listings')
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(listing_diff, 'Missing Listings')

            if not customer_diff.empty:
                # Provide an explanation
                customer_exp = 'Customer was found in Cleaning Fee Report but it was not found in Customer Report'
                customer_diff = explanation_missing(customer_diff, customer_exp)

                customer_diff.to_excel(writer, index=False, sheet_name='Missing Customer')
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(customer_diff, 'Missing Customer')

            if not man_issues.empty:
                # Provide an explanation
                management_exp = 'The Management Percent (Magpercent) is empty and there is no "omit" in the ' \
                                 'Management column.'
                man_issues = explanation_missing(man_issues, management_exp)

                man_issues.to_excel(writer, index=False, sheet_name='Missing Management')
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(man_issues, 'Missing Management')

            if not tax_issues.empty:
                # Provide an explanation
                tax_exp = 'The Tax Location (TaxLocation) is empty and there is no "omit" in the Hosp column.'
                tax_issues = explanation_missing(tax_issues, tax_exp)

                tax_issues.to_excel(writer, index=False, sheet_name='Missing Tax Location')
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(tax_issues, 'Missing Tax Location')

            if not vrbo_diff.empty:
                # Provide an explanation
                vrbo_exp = 'These VRBO entries did not have a customer attached to the ID, so their payouts were ' \
                           'not invoiced.'
                vrbo_diff = explanation_missing(vrbo_diff, vrbo_exp)

                vrbo_diff.to_excel(writer, index=False, sheet_name='Missing VRBO')
                # Note: It isn't possible to format any cells that already have a format such
                # as the index or headers or any cells that contain dates or datetimes.

                # Auto-adjust columns' width
                get_col_widths(tax_issues, 'Missing VRBO')

    return path, month_name
