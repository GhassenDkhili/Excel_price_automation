# Excel Price Automation

A simple Python project for processing Excel files and automatically calculating corrected prices using `openpyxl`.

## Features

- Reads data from an Excel file
- Processes product prices
- Applies a 10% price reduction
- Writes the corrected prices into a new column
- Saves the processed Excel file

## Technologies

- Python
- openpyxl
- Excel (.xlsx)
- Git & GitHub

## How It Works

The script reads prices from an Excel spreadsheet.  
Prices are converted from text values such as `$5.95` into numeric values.

A 10% reduction is then calculated:

`corrected_price = price * 0.9`

The corrected prices are written into a new column and the resulting workbook is saved as a new Excel file.

## Installation

Install the required Python package:

`pip install openpyxl`

## Run the Project

Run the Python script:

`python app.py`

## Project Structure

- `app.py` – Python script for processing the Excel file
- `transactions.xlsx` – Original Excel data
- `transactions_new.xlsx` – Processed Excel file
- `.gitignore` – Files and folders excluded from Git

