#!/usr/bin/env python3

import argparse
import numpy as np
import os
import pandas as pd
import re
import requests
import sys
import tabula
import warnings
import fitz

from collections import namedtuple
from datetime import datetime, timedelta
from pathlib import Path

TERMS_FILE = Path(__file__).parent / '.sap-stock-tool-accepted_terms'
DISCLAIMER_MESSAGE='''
This tool was developed for personal use to assist with organizing and preparing information for tax declaration purposes.

It is provided “as is” without any warranties or guarantees of accuracy, completeness, or fitness for a particular purpose.

While reasonable effort has been made to ensure correctness, you are solely responsible for reviewing, validating, and ensuring the accuracy of the information it generates.

Use of this tool does not replace your responsibility to comply with local tax laws or to seek professional tax advice when needed.

By using this tool, you agree that the author accepts no liability or responsibility for any errors, omissions, or consequences arising from its use.

Furthermore, this script is not affiliated with, endorsed by, or connected to SAP SE or any of its subsidiaries.

It was created independently by an SAP employee for personal use and is shared solely for the convenience of others.

Use at your own discretion and risk.
'''

VALUE_JSON_KEY="value"
COTACAO_COMPRA_JSON_KEY='cotacaoCompra'
COTACAO_VENDA_JSON_KEY='cotacaoVenda'
BCB_CACHE_FOR_EUR_REQUESTS={}

PDF_KEY_DATE_PURCHASED = 'Grant date'         # date of the purchase
PDF_KEY_COST_WHEN_PURCHASED = 'Cost basis /'  # actual value when purchased
PDF_KEY_QUANTITY_PURCHASED = 'Allocated /'    # quantity purchased

SHEET_KEY_DATE_WHEN_SOLD="Date"                  # date of the sell
SHEET_KEY_PRICE_WHEN_SOLD="Execution price"      # price of the asset when sold
SHEET_KEY_QUANTITY_SOLD="Quantity"               # quantity sold
SHEET_KEY_NET_PROCEEDS_WHEN_SOLD="Net proceeds"  # net proceeds (after any fees)

TAX_ON_PROFIT_PERCENTAGE=0.15  # 15%

from enum import Enum

class OpType(Enum):
    BUY = "BUY"
    SELL = "SELL"

class CurrencyConversionType(Enum):
    EUR_TO_BRL = "EUR_TO_BRL"
    BRL_TO_EUR = "BRL_TO_EUR"

class StockTransactionEntry:
    
    @staticmethod
    def sort_by_date(x):
        return (x.date, 0 if x.op_type == OpType.BUY else 1)

    def __init__(self, date, op_type, qty, total_qty, cur_conv_rate, cur_conv_type, avg_price,
                 price_eur=None, price_brl=None, net_proceeds=None, profit=None, tax=None):
        self.date = date
        self.op_type = op_type
        self.qty = qty
        self.total_qty = total_qty
        self.cur_conv_rate = cur_conv_rate
        self.cur_conv_type = cur_conv_type
        self.avg_price = avg_price
        self.price_eur = price_eur
        self.price_brl = price_brl
        self.net_proceeds = net_proceeds
        self.profit = profit
        self.tax = tax

    def __repr__(self):
        return (
            f"StockTransactionEntry(date={self.date}, op_type={self.op_type}, qty={self.qty}, "
            f"total_qty={self.total_qty}, cur_conv_rate={self.cur_conv_rate}, "
            f"cur_conv_type={self.cur_conv_type}, avg_price={self.avg_price}, "
            f"price_eur={self.price_eur}, price_brl={self.price_brl}, "
            f"net_proceeds={self.net_proceeds}, profit={self.profit})"
        )

    # expects ISO 8601 date format (YYYY-MM-DD)
    def get_date_as_tuple(self):
        return tuple(self.date.split('-'))



def enforce_terms_acceptance():
    if TERMS_FILE.exists():
        return True

    print("\n--- TERMS AND CONDITIONS ---\n")
    print(DISCLAIMER_MESSAGE)
    accept = input("Do you accept these terms and conditions? [y/N]: ").strip().lower()

    if accept == 'y':
        TERMS_FILE.write_text("Accepted")
        return True
    else:
        print("You must accept the terms to use this tool.")
        sys.exit(1)

# merge flow
def merge_data(input_files, output_path):
    save_stock_entries_to_excel(sorted((entry for in_file in input_files for entry in load_stock_entries_from_excel(in_file)), key=StockTransactionEntry.sort_by_date), output_path)

# process flow
def process_data(print_output, reverse, input_path, output_path):
    unprocessed_transactional_data = load_stock_entries_from_excel(input_path)
    
    processed_transactional_data = process_transactions(unprocessed_transactional_data)
    
    save_stock_entries_to_excel(processed_transactional_data, output_path)

    if print_output:
        print_transaction_list_as_table(processed_transactional_data, reverse)

def process_transactions(registry_of_transactions):
    total_qty = 0  # total quantity of current holdings
    total_cost_brl = 0  # total cost of current holdings
    current_avg_price = 0  #  total_cost_brl / total_qty

    for transaction in registry_of_transactions:
        if transaction.op_type == OpType.BUY:
            # calculate values
            buying_price_eur = get_price_for_buying_eur_at_date(*transaction.get_date_as_tuple())
            total_qty += transaction.qty
            price_brl = transaction.price_eur * transaction.qty * buying_price_eur
            total_cost_brl += price_brl
            current_avg_price = total_cost_brl / total_qty

            # update transaction record
            transaction.cur_conv_rate = buying_price_eur
            transaction.total_qty = total_qty
            transaction.price_brl = price_brl
            transaction.total_cost_brl = total_cost_brl
            transaction.avg_price = current_avg_price

        elif transaction.op_type == OpType.SELL:
            # calculate values
            selling_price_eur = get_price_for_selling_eur_at_date(*transaction.get_date_as_tuple())
            total_qty -= transaction.qty
            total_cost_brl -= transaction.qty * current_avg_price

            # calculate profit based on average price
            profit = selling_price_eur * transaction.net_proceeds - transaction.qty * current_avg_price

            # update transaction record
            transaction.cur_conv_rate = selling_price_eur
            transaction.total_qty = total_qty
            transaction.profit = profit
            transaction.avg_price = current_avg_price
            transaction.price_brl = transaction.price_eur * selling_price_eur
            transaction.tax = profit * TAX_ON_PROFIT_PERCENTAGE
        else:
            raise ValueError(f"Unsupported Enum in OpType: {transaction.op_type}")

    return registry_of_transactions

def get_price_for_buying_eur_at_date(year, month, day):
    return float(_get_eur_quotation_data_for_date(year, month, day)[VALUE_JSON_KEY][0][COTACAO_COMPRA_JSON_KEY])

def get_price_for_selling_eur_at_date(year, month, day):
    return float(_get_eur_quotation_data_for_date(year, month, day)[VALUE_JSON_KEY][0][COTACAO_VENDA_JSON_KEY])

def yesterday_str(date_tuple):
    year_s, month_s, day_s = date_tuple
    # build a datetime object (time portion is ignored)
    dt = datetime(int(year_s), int(month_s), int(day_s))
    # subtract one day
    prev = dt - timedelta(days=1)
    # format back to strings
    return (
        f"{prev.year:04d}",
        f"{prev.month:02d}",
        f"{prev.day:02d}"
    )

def _get_eur_quotation_data_for_date(year, month, day):
    date = f'{month}-{day}-{year}' # for some stupid reason it uses 'murica weird date format..

    result = BCB_CACHE_FOR_EUR_REQUESTS.get(date, None)  # default = None

    done = False

    if result is None:
        while not done:
            print("Trying ", date, "...")
            result = requests.get(f"https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoMoedaPeriodoFechamento(codigoMoeda=@codigoMoeda,dataInicialCotacao=@dataInicialCotacao,dataFinalCotacao=@dataFinalCotacao)?@codigoMoeda='EUR'&@dataInicialCotacao='{date}'&@dataFinalCotacao='{date}'&$format=json").json()
            if len(result['value']) == 0:
                year, month, day = yesterday_str((year, month, day))
                new_date = f'{month}-{day}-{year}' # for some stupid reason it uses 'murica weird date format..
                print("Day is not valid (weekend or holiday); changed from ", date, " to ", new_date)
                date = new_date
            else:
                done = True
            
            print(result)
            BCB_CACHE_FOR_EUR_REQUESTS[date] = result

    return result

# extract flow
def extract_data(fiscal_year, buy_data_path, sell_data_path, output_path):
    buy_data = extract_buy_data_from_pdf(buy_data_path, fiscal_year)
    sell_data = extract_sell_data_from_spreadsheet(sell_data_path, fiscal_year)
    unprocessed_transactional_data = merge_transactional_data(buy_data, sell_data)
    save_stock_entries_to_excel(unprocessed_transactional_data, output_path)

def extract_until_text(pdf_path, search_text):
    # Step 1: Search for the text in raw PDF content
    doc = fitz.open(pdf_path)
    target_page = None

    for page_num in range(len(doc)):
        text = doc[page_num].get_text().lower()
        if search_text.lower() in text:
            target_page = page_num + 1 # Pages are 1-based
            break

    doc.close()

    # Step 2: If found, extract tables only up to that page
    if target_page is not None:
        print(f"Stopping at page {target_page}")
        dfs = tabula.read_pdf(pdf_path, pages=f"1-{target_page - 1}")
    else:
        print(f"Extracting all pages.")
        dfs = tabula.read_pdf(pdf_path, pages='all')

    return dfs

def extract_buy_data_from_pdf(path, fiscal_year):
    pattern = rf'\d{{1,2}} [A-Z][a-z]{{2}} {fiscal_year}'
    is_fiscal_year_entry = lambda x: type(x) is str and re.fullmatch(pattern, x) is not None

    dfs = extract_until_text(path, "Portfolio 2 - Positions - Restricted shares")

    for df in dfs:
        df.columns = df.iloc[0] # set the headers with something that actually makes sense

    _dfs = [df for df in dfs if PDF_KEY_DATE_PURCHASED in df.columns]
    valid_df = pd.concat(_dfs, ignore_index=True)
    filtered_df = valid_df[valid_df[PDF_KEY_DATE_PURCHASED].apply(is_fiscal_year_entry)]
    
    filtered_df[PDF_KEY_DATE_PURCHASED] = pd.to_datetime(filtered_df[PDF_KEY_DATE_PURCHASED], format='%d %b %Y', errors='coerce')
    filtered_df[PDF_KEY_DATE_PURCHASED] = filtered_df[PDF_KEY_DATE_PURCHASED].dt.strftime('%Y-%m-%d')
    
    filtered_df[PDF_KEY_COST_WHEN_PURCHASED] = np.where(
        filtered_df[PDF_KEY_COST_WHEN_PURCHASED].str.endswith(" EUR"),
        filtered_df[PDF_KEY_COST_WHEN_PURCHASED].str.slice(0, -4),
        filtered_df[PDF_KEY_COST_WHEN_PURCHASED]
    ).astype(float)

    filtered_df[PDF_KEY_QUANTITY_PURCHASED] = filtered_df[PDF_KEY_QUANTITY_PURCHASED].astype(float)

    return zip(
            filtered_df[PDF_KEY_DATE_PURCHASED], # date of the purchase
            filtered_df[PDF_KEY_COST_WHEN_PURCHASED], # actual value when purchased
            filtered_df[PDF_KEY_QUANTITY_PURCHASED]  # quantity purchased
        )

def extract_sell_data_from_spreadsheet(file_path, year):
    # Load the Excel sheet, skipping the first 4 rows (headers start at row 5)
    df = pd.read_excel(file_path, skiprows=4)
    is_fiscal_year_entry = lambda date: date.year == int(year)

    # Filter rows where Order type is "Sell at market price" or "Sell with price limit" and Status is "Executed"
    filtered_df = df[
        df["Order type"].astype(str).isin(["Sell at market price", "Sell with price limit", "Sell-to-cover", "Sell"]) &
        (df["Status"] == "Executed") &
        (df["Date"].apply(is_fiscal_year_entry)) &
        (df["Product type"] == "shares")
    ]

    filtered_df[SHEET_KEY_DATE_WHEN_SOLD] = filtered_df[SHEET_KEY_DATE_WHEN_SOLD].dt.strftime('%Y-%m-%d')

    return zip(
            filtered_df[SHEET_KEY_DATE_WHEN_SOLD],
            filtered_df[SHEET_KEY_PRICE_WHEN_SOLD],
            filtered_df[SHEET_KEY_QUANTITY_SOLD],
            filtered_df[SHEET_KEY_NET_PROCEEDS_WHEN_SOLD]
        )

def merge_transactional_data(buy_transactional_data, sell_transactional_data, sort=True):
    registry_of_transactions = []

    for date, price_eur, qty in buy_transactional_data:
        registry_of_transactions.append(StockTransactionEntry(date, OpType.BUY, qty, None, None, CurrencyConversionType.EUR_TO_BRL, None, price_eur=price_eur))

    for date, execution_price, qty, net_proceeds in sell_transactional_data:
        registry_of_transactions.append(StockTransactionEntry(date, OpType.SELL, qty, None, None, CurrencyConversionType.BRL_TO_EUR, None, price_eur=execution_price, net_proceeds=net_proceeds))

    if sort:
        registry_of_transactions.sort(key=StockTransactionEntry.sort_by_date)

    return registry_of_transactions

# io-related
def format_value(val, fmt="{:<16}", none_placeholder="{:<16}".format("—")):
    return none_placeholder if pd.isna(val) else fmt.format(val)

def print_transaction_list_as_table(list_of_transactions, reverse=False):
    # Header
    header = [
        "Date", "Op Type", "Quantity", "Total Quantity", "€ Conv. Rate", 
        "Conv. Type", "Avg Price (EUR)", "Price (EUR)", "Price (BRL)", 
        "Net Proceeds", "Profit (BRL)", "Tax (BRL)"
    ]
    table = " | ".join(f"{h:<16}" if h != "Op Type" else f"{h:<8}" for h in header) + "\n"
    table += "-" * len(table) + "\n"

    printed_list = reversed(list_of_transactions) if reverse else list_of_transactions
    
    # Rows
    for entry in printed_list:
        table += " | ".join([
            f"{entry.date:<16}",
            f"{entry.op_type.name:<8}",
            f"{entry.qty:<16.5f}",
            f"{entry.total_qty:<16.5f}",
            format_value(entry.cur_conv_rate, "{:<16.4f}"),
            format_value(entry.cur_conv_type.name),
            format_value(entry.avg_price, "{:<16.2f}"),
            format_value(entry.price_eur, "{:<16.2f}"),
            format_value(entry.price_brl, "{:<16.2f}"),
            format_value(entry.net_proceeds, "{:<16.2f}"),
            format_value(entry.profit, "{:<16.2f}"),
            format_value(entry.tax, "{:<16.2f}")
        ]) + "\n"

    print(table)

def save_stock_entries_to_excel(entries, output_path):
    # Suppose `entries` is your list of StockTransactionEntry instances
    entries_data = [entry.__dict__ for entry in entries]

    # Convert to DataFrame
    df = pd.DataFrame(entries_data)

    # Write to Excel
    df.to_excel(output_path, index=False)

def parse_enum(enum_class, value):
    """
    Parses a string like 'OpType.BUY' into enum_class.BUY,
    ensuring that the enum class name matches.

    Args:
        enum_class: The Enum class to use (e.g., OpType).
        value: The string (e.g., 'OpType.BUY').

    Returns:
        The enum value (e.g., OpType.BUY), or None if not matched.
    """
    if not value or not isinstance(value, str):
        return None

    enum_name = enum_class.__name__
    pattern = rf'^{enum_name}\.(\w+)$'
    match = re.match(pattern, value)

    if match:
        member_name = match.group(1)
        try:
            return enum_class[member_name]
        except KeyError:
            return None  # Member not found
    return None

def load_stock_entries_from_excel(input_path):
    df = pd.read_excel(input_path)

    entries = []
    for _, row in df.iterrows():
        entry = StockTransactionEntry(
            date=row['date'],
            op_type=parse_enum(OpType, row['op_type']),
            qty=row['qty'],
            total_qty=row['total_qty'],
            cur_conv_rate=row['cur_conv_rate'],
            cur_conv_type=parse_enum(CurrencyConversionType, row['cur_conv_type']),
            avg_price=row['avg_price'],
            price_eur=row.get('price_eur'),
            price_brl=row.get('price_brl'),
            net_proceeds=row.get('net_proceeds'),
            profit=row.get('profit'),
            tax=row.get('tax')
        )
        entries.append(entry)
    
    return entries

# parser
def parse_arguments():
    parser = argparse.ArgumentParser(
        description="SAP Stock Tool - Extract and Process Equity Data"
    )

    # Top-level argument for showing warnings
    parser.add_argument(
        '-w', '--show-warnings',
        action='store_true',
        help='Show warnings (default: suppress warnings)',
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # Extract subcommand
    extract_parser = subparsers.add_parser(
        "extract", help="Extracts equity data from EquatePlus files"
    )
    extract_parser.add_argument(
        "-y", "--year", type=int, required=True,
        help="Fiscal year for the report"
    )
    extract_parser.add_argument(
        "-b", "--buy-data", type=str, required=True,
        help="Path to the PDF EoY buy data file from EquatePlus"
    )
    extract_parser.add_argument(
        "-s", "--sell-data", type=str, required=True,
        help="Path to the Excel file with transaction history"
    )
    extract_parser.add_argument(
        "-o", "--out", type=str, default="output-unprocessed.xlsx",
        help="Path to output Excel file for aggregated data"
    )

    # Merge subcommand
    merge_parser = subparsers.add_parser(
        "merge", help="Merge multiple extracted Excel files into one"
    )

    merge_parser.add_argument(
        "-i", "--in", dest="input_files", nargs="+", required=True,
        help="List of input Excel files to merge"
    )

    merge_parser.add_argument(
        "-o", "--out", type=str, default="output-unprocessed.xlsx",
        help="Path to output merged Excel file, by default it will write to output-unprocessed.xlsx"
    )

    # Process subcommand
    process_parser = subparsers.add_parser(
        "process", help="Processes previously extracted data"
    )
    process_parser.add_argument(
        "-p", action="store_true",
        help="Print results to terminal"
    )
    process_parser.add_argument(
        "-r", action="store_true",
        help="Reverse the order of terminal output (latest first)"
    )
    process_parser.add_argument(
        "-i", "--in", dest="input_path", type=str, default="output-unprocessed.xlsx",
        help="Path to the extracted output Excel file (defaults to output-unprocessed.xlsx)"
    )
    process_parser.add_argument(
        "-o", "--out", dest="output_path", type=str, default="output-processed.xlsx",
        help="Path to save processed output (defaults to output-processed.xlsx)"
    )

    return parser.parse_args()

def prevent_overwrite(out_path):
    if os.path.exists(out_path):
        print(f"file {out_path} already exists. Please provide an empty path for output.")
        sys.exit(1)

def main():
    args = parse_arguments()

    if not args.show_warnings:
        pd.options.mode.chained_assignment = None  # default='warn'
        warnings.filterwarnings("ignore", message=".*Failed to import jpype dependencies.*")

    ensure_xlsx_extension = lambda path: path if path.endswith(".xlsx") else path + ".xlsx"

    if args.command == "extract":
        out_path = ensure_xlsx_extension(args.out)
        prevent_overwrite(out_path)
        extract_data(args.year, args.buy_data, args.sell_data, out_path)
    elif args.command == "merge":
        out_path = ensure_xlsx_extension(args.out)
        merge_data(args.input_files, out_path)
    elif args.command == "process":
        out_path = ensure_xlsx_extension(args.output_path)
        prevent_overwrite(out_path)
        process_data(args.p, args.r, args.input_path, out_path)
    else:
        parser.print_help()
        sys.exit(1)

if __name__ == "__main__":
    enforce_terms_acceptance()
    main()
