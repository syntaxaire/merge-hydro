# Process Hydro sheets to meter readings

import os
import re
import xlrd
import logging
import pandas as pd
import configparser
import numpy as np

logging.basicConfig(level=logging.DEBUG, format="%(levelname)s:%(message)s")

cwd = os.getcwd()
print(f"Current working directory is {cwd}")


def find_in_sheet(val, sheet: xlrd.sheet.Sheet):
    """Return a tuple containing the (row, col) of first match searching row 0, then row 1, etc.
    Return false if not found."""
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if sheet.cell_value(row, col) == val:
                return row, col
    return False


def get_billing_lines(sheet: xlrd.sheet.Sheet):
    """Take an xl sheet with a Hydro One bill as input and return the number of billing lines (accounts)"""
    lines = 0
    row, col = find_in_sheet("Line #", sheet)  # this cell should be unique

    # find number of lines in bill by walking down
    row += 1
    while sheet.cell_type(row, col) == 2:  # Excel float cell type
        row += 1
        lines += 1
    return lines


def get_bill(filename: str) -> pd.DataFrame:
    """Take the name of an Excel file containing a Hydro One bill and return a Pandas dataframe
    constructed from a select area of the invoice sheet."""

    logging.debug(f"Reading {filename}")
    with xlrd.open_workbook(filename=filename) as xl:
        try:
            sheet = xl.sheet_by_name("Invoice Summary")  # this sheet contains the bill
        except:
            logging.error(f"Not able to open sheet \"Invoice Summary\" in {filename}")
            raise

        num_rows = sheet.nrows
        lines = get_billing_lines(sheet)
        logging.info(f"{filename} has {lines} account lines.")
        label_row, _ = find_in_sheet("Line #", sheet)  # index of row with column headers
        logging.debug(f"Label row index: {label_row}")
        _, last_col = find_in_sheet("Metered Usage [kWh]", sheet)  # index of last desired column
        logging.debug(f"Last column to parse: {last_col}")
        footer_size = num_rows - (label_row + 1) - lines
        logging.debug(f"Footer size: {footer_size}")

        return pd.read_excel(io=filename,
                             sheet_name="Invoice Summary",
                             header=label_row,
                             skip_footer=footer_size,
                             index_col=0,
                             usecols=list(range(1, last_col + 1)))


config_file = "process.cfg"
config = configparser.ConfigParser(allow_no_value=True)
config.read(config_file)

using_aliases = False
if "Aliases" in config:
    using_aliases = True
    aliases = {}
    for account in config["Aliases"]:
        aliases[account] = config["Aliases"][account]
    logging.debug(f"Account aliases: {aliases}")

spreadsheets = []
files = os.listdir(path=".")

for file in files:
    if re.match("^[0-9]{8}.xls", file):
        spreadsheets.append(file)
logging.debug(f"Files to process: {spreadsheets}")
bills_list = [get_bill(filename) for filename in spreadsheets]
mass_df = pd.concat(bills_list)  # join all bills in preparation for splitting by account number
mass_df['Reading From Date'] = pd.to_datetime(mass_df['Reading From Date'])  # convert date columns from strings
mass_df['Reading From Date'] = [dt.to_datetime().date() for dt in mass_df['Reading From Date']]
mass_df['Reading To Date'] = pd.to_datetime(mass_df['Reading To Date'])
mass_df['Reading To Date'] = [dt.to_datetime().date() for dt in mass_df['Reading To Date']]
mass_df['Days In Reading'] = mass_df['Reading To Date'] - mass_df['Reading From Date']
mass_df['Days In Reading'] = mass_df['Days In Reading'] / np.timedelta64(1, 'D')
mass_df['kWh Per Day'] = mass_df['Metered Usage [kWh]'] / mass_df['Days In Reading']

# drop certain types of rows: unbilled entries, and sentinel lights
mass_df = mass_df[mass_df["Service Classification"] != "Sentinel Lights"]
mass_df = mass_df[mass_df["Reason Not Billed"] != "No billing as of summary billing cut off date"]

# after processing, drop columns according to config
if "Drop" in config:
    for drop_label in mass_df.columns.values.tolist():
        if drop_label.lower() in config["Drop"]:
            mass_df = mass_df.drop(labels=drop_label, axis=1)
            logging.info(f"Dropped column \"{drop_label}\" according to config")

accounts = set(mass_df.index)  # unique account numbers

# write a spreadsheet with processed results
with pd.ExcelWriter("output.xlsx") as writer:
    for account_num in accounts:
        if using_aliases:
            worksheet_name = aliases[str(account_num)]
        else:
            worksheet_name = str(account_num)
        mass_df.loc[account_num].to_excel(writer, worksheet_name)
        logging.info(f"Wrote results for account {account_num} to output.xlsx in sheet {worksheet_name}")
