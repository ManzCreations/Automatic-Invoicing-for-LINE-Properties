#######################################################################################################################
# Modules
import datetime
import os
from collections import defaultdict
from datetime import datetime

import numpy as np

from helpful_tools import *

# from unidecode import unidecode

# Disable copy warnings
pd.options.mode.chained_assignment = None


#######################################################################################################################
# Notes
# Check check and CM dates, make sure that they are set to the first of the month.


def line_invoice_generation(app):
    app.progress_bar["value"] = 1
    app.progress_bar.update()

    ####################################################################################################################
    # Setup and initialization
    filenames = ['reservations', 'airbnb', 'Current', 'VRBO_']
    month, year = date_from_airbnb_name(filenames[1])
    invoice_date, due_date = generate_dates(month, year)
    month_name = month_number_to_name(month)

    app.progress_bar["value"] = 2
    app.progress_bar.update()

    path_month = setup_directory_structure(app, month_name, year)
    filepath = os.path.join(os.getcwd(), 'ModelFiles')
    dataframes, reformat_info = load_files(app, filepath, filenames, path_month)

    # Data cleaning and reformatting
    customer_info_df = dataframes.get('customer_info', pd.DataFrame())
    for key, df in dataframes.items():
        if key in ['bnb', 'check', 'vrbo']:  # Apply reformatting to these specific dataframes
            dataframes[key] = reformat_and_update_files(df, customer_info_df)

    dataframes = clean_dataframes(dataframes)

    ####################################################################################################################
    # Create data frames of information
    app.log("Optimizing Data...")
    app.progress_bar["value"] = 16
    app.progress_bar.update()

    replace_re = "[^A-Za-z0-9_ -:&]+"
    # Assuming bnb, cleaning, customer_info, check, and vrbo are previously defined DataFrames
    bnb_col = prepare_dataframe_columns(bnb, ['Listing', 'Amount', 'Type', 'Confirmation Code', 'Nights'], replace_re)
    cleaning_col = prepare_dataframe_columns(cleaning,
                                             ['ListingBNB', 'QBO', 'Cleaning', 'Tax_Location', 'Pest', 'Landscape',
                                              'Internet/Cable', 'Bus_Lic', 'VRBO_ID', 'Code', 'Output'], replace_re)
    customer_col = prepare_dataframe_columns(customer_info,
                                             ['Customer-QBO', 'Expense_Flat', 'Credit', 'Clean', 'Hosp', 'Management',
                                              'Magpercent'], replace_re)
    check_col = prepare_dataframe_columns(check, ['Listing'], replace_re)
    # vrbo_col preparation would follow a similar pattern

    # VRBO data management - placeholder functions for mtn and number2month need to be defined or imported
    vrbo_save_path = os.path.join(filepath, 'VRBO Date Organizer.xlsx')
    manage_vrbo_data(vrbo, vrbo_save_path, datetime.now().month, datetime.now().strftime("%B"), datetime.now().year,
                     month_name_to_number, month_number_to_name)

    # Copy VRBO Excel file to report folder
    copy_excel_file(vrbo_save_path, os.path.join(path, 'VRBO Date Organizer.xlsx'))
    ###################################################################################################################
    # Missing information and general housekeeping
    app.log("Finding Missing Information...")

    # Assuming bnb_col, cleaning_col, check_col are defined DataFrames
    bnb_col = remove_extra_spaces(bnb_col, ['Listing'])
    cleaning_col = remove_extra_spaces(cleaning_col, ['ListingBNB'])
    check_col = remove_extra_spaces(check_col, ['Listing'])

    listing_diff = find_diff_and_concat(bnb_col, cleaning_col, 'Listing', 'ListingBNB')
    vrbo_diff = find_diff_and_concat(vrbo_col, cleaning_col, 'Property ID', 'VRBO_ID')

    # Assuming customer_diff_qbo logic and modifications are done elsewhere based on the context

    app.log("Separating Aviad's Listings...")
    separate_listings_based_on_output(cleaning_col, bnb, vrbo, path, month, year)
    ###################################################################################################################
    app.log("Creating total unit containing all necessary data...")
    unit = aggregate_customer_data(customer_col, cleaning_col, bnb_col, check_col, vrbo_col, month)

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
                                    [invoice_no, unit_repeat['Customer'].iloc[0], invoice_date, due_date, '', item[0],
                                     item[0],
                                     num_cleaning_fee,
                                     round(cleaning_fee, 2), round(cleaning_fee * num_cleaning_fee, 2), tax,
                                     invoice_date])
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
                        [invoice_no, customer_name, invoice_date, due_date, memo, item[2], item[2], 0.01 * man_rate,
                         round(total_management, 2), round(management_fee, 2), tax_location, invoice_date])
                    entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Log Pest Control Fee
            if not unit_loop['Pest'].isnull().values.all():
                pest_fee = unit_loop['Pest'].sum()
                entry_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice_date, due_date, memo, 'SERVICES', 'PEST CONTROL', 1,
                     round(pest_fee, 2), round(pest_fee, 2), tax_location, invoice_date])
                entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Log Landscaping Fee
            if not unit_loop['Landscape'].isnull().values.all():
                landscaping_fee = unit_loop['Landscape'].sum()
                entry_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice_date, due_date, memo, 'SERVICES', 'LANDSCAPING', 1,
                     round(landscaping_fee, 2), round(landscaping_fee, 2), tax_location, invoice_date])
                entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Log Internet/Cable Fee
            if not unit_loop['Internet/Cable'].isnull().values.all():
                cable_fee = unit_loop['Internet/Cable'].sum()
                entry_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice_date, due_date, memo, 'SERVICES', 'INTERNET_CABLE', 1,
                     round(cable_fee, 2), round(cable_fee, 2), tax_location, invoice_date])
                entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Log Pest Control Fee
            if not unit_loop['Bus_Lic'].isnull().values.all():
                bus_lic = unit_loop['Bus_Lic'].sum()
                entry_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice_date, due_date, memo, 'SERVICES', 'BUSINESS LICENSE', 1,
                     round(bus_lic, 2), round(bus_lic, 2), tax_location, invoice_date])
                entry = pd.concat([entry, entry_new.transpose()], ignore_index=True)

            # Credit Memo/Checks
            if (unit_loop['CreditMemo'] == 'CM').any():
                # Create item description for Credit Memo
                item_cm = 'Deposits in Trust'
                item_cm_desc = 'Funds Collected on Behalf of Client'

                # Log Credit Memo
                credit_memo_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice_date, invoice_date, item_cm, item_cm_desc,
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
                line_itm = 'Deposits in Trust'

                # Log Sales Receipts
                sr_new = pd.DataFrame(
                    [invoice_no, customer_name, invoice_date, tax_location, item_sr, pmnt_meth,
                     sales_msg, 'N', 'N', invoice_date, line_itm, sales_msg, 1, total_sale_amt, total_sale_amt, 'NON'])
                sales_receipts = pd.concat([sales_receipts, sr_new.transpose()], ignore_index=True)

                # Checks
                # Reference number
                ref_num = 'ABB TR ' + f'{check_no:05d}'
                # Bank Account
                bank = 'ABB Trust #5241'

                # Expense descriptions
                expense_desc_object = datetime.strptime(str(month), '%m')
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
                    [ref_num, bank, invoice_date, customer_name, expense_amt, private_note, expense_desc,
                     expense_acc])
                checks = pd.concat([checks, checks_new.transpose()], ignore_index=True)
                check_no += 1

                # Journal Entries
                ref_num = 'PMT ' + f'{journal_no:05d}'
                customer = unit_loop['Customer'].iloc[0]
                prvt_note = f'Payment {customer} to #{invoice_no}'
                debit_acct = 'Accounts Receivable (A/R)'
                credit_acct = 'ABB Trust #5241'

                # Log Journal Entries
                # Debit
                je_new = pd.DataFrame(
                    [ref_num, invoice_date, prvt_note, 'False', debit_acct, total_invoice, debit_acct, tax_location,
                     customer])
                journal_entries = pd.concat([journal_entries, je_new.transpose()], ignore_index=True)
                # Credit
                je_new = pd.DataFrame(
                    [ref_num, invoice_date, prvt_note, 'False', credit_acct, -total_invoice, credit_acct, tax_location,
                     customer])
                journal_entries = pd.concat([journal_entries, je_new.transpose()], ignore_index=True)
                journal_no += 1

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
                               'LineDesc', 'Location', 'Entity']

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
