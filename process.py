# Process Hydro sheets to meter readings

import configparser
import logging
import xlrd
from pathlib import Path
from typing import Generator, Tuple

import numpy as np
import pandas as pd

logging.basicConfig(level=logging.DEBUG, format="%(levelname)s:%(message)s")

XLS_FLOAT_TYPE = 2

converters = {"Account Number": int}


def find_in_sheet(val, sheet: xlrd.sheet.Sheet) -> Tuple[int, int]:
    """Return a tuple containing the (row, col) of first match searching row 0, then row 1, etc."""
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if sheet.cell_value(row, col) == val:
                return row, col
    raise LookupError(f'Value {val} not found in sheet {sheet}')


def get_billing_lines(sheet: xlrd.sheet.Sheet) -> int:
    """Take an xl sheet with a Hydro One bill as input and return the number of billing lines (accounts)"""
    lines = 0
    try:
        row, col = find_in_sheet("Line #", sheet)  # this cell should be unique
    except LookupError:
        raise ValueError(f'{sheet} does not appear to be a Hydro One bill (no "Line #" cell)')

    # find number of lines in bill by walking down
    row += 1
    while sheet.cell_type(row, col) == XLS_FLOAT_TYPE:
        row += 1
        lines += 1
    return lines


def get_bill_dataframe(filename: Path) -> pd.DataFrame:
    """Take the name of an Excel file containing a Hydro One bill and return a Pandas dataframe
    constructed from a select area of the invoice sheet."""

    logging.debug(f"Reading {filename}")
    with xlrd.open_workbook(filename=filename) as xl:
        try:
            sheet = xl.sheet_by_name("Invoice Summary")  # this sheet contains the bill
        except xlrd.biffh.XLRDError:
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
                             skipfooter=footer_size,
                             index_col=0,
                             converters=converters,
                             usecols=list(range(1, last_col + 1)))


def process(spreadsheets: Generator[Path, None, None], config) -> pd.DataFrame:
    bills = [get_bill_dataframe(filename) for filename in spreadsheets]
    logging.debug("Completed loading dataframes from Excel.")
    mass_df = pd.concat(bills)  # join all bills in preparation for splitting by account number
    # convert date columns from strings
    mass_df['Reading From Date'] = pd.to_datetime(mass_df['Reading From Date'])
    mass_df['Reading From Date'] = [dt.date() for dt in mass_df['Reading From Date']]
    mass_df['Reading To Date'] = pd.to_datetime(mass_df['Reading To Date'])
    mass_df['Reading To Date'] = [dt.date() for dt in mass_df['Reading To Date']]

    # new columns
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

    return mass_df


def write_output(mass_df: pd.DataFrame, aliases: dict):
    accounts = set(mass_df.index)  # unique account numbers

    # write a spreadsheet with processed results
    with pd.ExcelWriter("output.xlsx") as writer:
        for account_num in accounts:
            if str(account_num) in aliases:
                worksheet_name = aliases[str(account_num)]
            else:
                worksheet_name = str(account_num)
            mass_df.loc[account_num].to_excel(writer, worksheet_name)
            logging.info(f"Wrote results for account {account_num} to output.xlsx in sheet {worksheet_name}")


def main():
    cwd = Path.cwd()
    config_file = cwd / "process.cfg"
    if not config_file.is_file():
        print(f"{config_file} not found. Please copy the example config to this file and edit it.")
    config = configparser.ConfigParser(allow_no_value=True)
    config.read(config_file)
    if "Aliases" in config:
        aliases = dict(config["Aliases"])
    else:
        aliases = {}
    logging.debug(f"Account aliases: {aliases}")
    spreadsheets = cwd.glob("*.xls")
    logging.debug(f"Files to process: {spreadsheets}")
    mass_df = process(spreadsheets, config)
    write_output(mass_df, aliases)


if __name__ == "__main__":
    main()
