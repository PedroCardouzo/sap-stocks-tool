# SAP Stocks Tool

## Disclaimer (VERY IMPORTANT!)

This tool was developed for personal use to assist with organizing and preparing information for tax declaration purposes.

It is provided “as is” without any warranties or guarantees of accuracy, completeness, or fitness for a particular purpose.

While reasonable effort has been made to ensure correctness, you are solely responsible for reviewing, validating, and ensuring the accuracy of the information it generates.

Use of this tool does not replace your responsibility to comply with local tax laws or to seek professional tax advice when needed.

By using this tool, you agree that the author accepts no liability or responsibility for any errors, omissions, or consequences arising from its use.

Furthermore, this script is not affiliated with, endorsed by, or connected to SAP SE or any of its subsidiaries.

It was created independently by an SAP employee for personal use and is shared solely for the convenience of others.

Use at your own discretion and risk.

## Setup

`pip install pandas requests tabula-py openpyxl numpy`
A relatively new python 3 version to ensure compatibility, although some care has been taken to ensure compatibility with older versions, I run it on 3.9.

## Usage

The tool works in two steps: extraction and processing. It will always ask for your acceptance of the limited liability and create a file in the same folder as the script `.sap-stock-tool-accepted_terms` so you don't have to keep accepting the terms every use.


### Extraction

Process the PDF File from EquatePlus End of Year statement, plus the Excel file that contains all the transactions into a format that can be enhanced, and processed later.

```bash
python sap-stock-tool.py extract -y <fiscal-year> -b <buy-pdf-path> -s <sell-excel-path> -o <output-file>
```

#### Arguments

* `-y`, `--year` (required): Fiscal year to extract data for (e.g., `2024`).
* `-b`, `--buy-data` (required): Path to the EquatePlus PDF file for end-of-year stock buys. It must be the one containing the relevant data for the fiscal year provided at the `-y` option.
* `-s`, `--sell-data` (required): Path to the Excel file from EquatePlus with transaction history. This includes selling, dividends received, and RSUs. This program is focusing on only executed sell at market price, or sell at limited value operations, and is ignoring RSUs completely for now, as well as dividends (SAP pays them 1/year, just copy the data from the spreadsheet directly).
* `-o`, `--out`: Path to the output Excel file for the extracted data (default: `output-unprocessed.xlsx`). There will be a lot of empty fields that will be filled later in the "Processing" step. If the file already exists, it will throw an error to prevent unwanted overwrites. If you want to compute for multiple years, you can run this once per file, then merge the files in order, with a single header. (later we will have a merge functionality which is currently WIP)

You can now make sure the data is correct by opening the output file, and can make any manual adjustments as needed, which will be considered in the processing, like changing the date you sold something because you actually converted from EUR to BRL in another day with very different exchange rates, or include some other entry that the extraction ignored.

### Merge

Merge multiple extracted Excel files into a single combined dataset. This is useful if you've run the extraction step for multiple fiscal years or across different sources and want to consolidate them before processing.

```bash
sap-stock-tool.py merge -in <file1> <file2> ... <fileN> -out <output_file>
```

#### Arguments

* `-i`, `--in`: List of paths to one or more previously extracted `.xlsx` files to merge.
* `-o`, `--out`: Path to the output file that will store the merged dataset in Excel format (default: `output-unprocessed.xlsx`).

### Processing

Processes the previously extracted Excel file, computing average prices, BRL conversions, and profit/loss.

```bash
python sap-stock-tool.py process -p -r -i <input-file> -o <output-file>
```

#### Arguments

* `-p`: Prints results to the terminal.
* `-r`: Reverses terminal output (first lines will be the latest transactions in the terminal output only).
* `-i`, `--in`: Path to the extracted Excel file (default: `output-unprocessed.xlsx`). Remember to review it and adjusted manually for anything that was outside expectations.
* `-o`, `--out`: Path to save processed data (default: `output-processed.xlsx`). Will output a similar file, but with the fields properly filled. If the file already exists, it will throw an error to prevent unwanted overwrites.

`sap-stock-tool.py process -pr -in <path/to/compiled/data> -out <path/to/output> `
It will go from oldest (first record in the provided excel) to newest (last record in the provided sheet), cross-referencing the price with the Banco Central data to get the EUR PTAX, and using the net proceeds from EquatePlus data as well as calculating average price - will arrive at your profits and will have a field with the due amount to pay in tax. You can then add up all the items in the current year and arrive at what you should expect to see being charged by the Receita Federal in the tax declaration app.

## To Do
 - Fix Excel output for the process command is including a "total_cost_brl" column.. idk where it comes from...
 - Add guide with images on how to get the data from EquatePlus
