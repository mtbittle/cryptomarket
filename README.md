# Crypto CoinMarket Data Puller


## About
Uses CoinMarket Pro API to fetch latest prices for specific crypto. This is very rudimentary and does not contain
much error handling (or likely python best practices - I'm green with python). The script assumes A LOT:
- The script, config file, and excel file all exist in the same directory
- The Excel file name matches what is in the config.properties file
- All properties in the `config.properties` file have values
- The crypto symbols you list in the props file are real (and likely have to be uppercase)
- The crypto symbols and the excel mappings in the props file are relative:
  - Example below would mean the BTC price would be updated in excel cell A1, and ETH in cell B2
    - crypto.symbols=BTC,ETH
    - excel.mapping=A1,B2
- The Excel file is NOT open when you run the script
- Whew...that might cover everything

# Setup / Configuration
1. Check if you have Python 3 installed (via cmd window): `py --version`
2. Install Python 3 if needed. For Windows, it is likely easiest to use the recommended 'Installer' package.
Go to [Download Python](https://www.python.org/downloads/windows), select the latest release link, then scroll down to the Windows Installer (64-bit).
3. To write to Excel files, you need to install the `openpyxl` package via `pip`:
   1. open a command line window and enter: `py -m pip install openpyxl`
4. Go to [CoinMarketCap API Page](https://coinmarketcap.com/api/) and sign up for the free basic plan.
There is a limit to how many times you can call the API per day, so choose a different plan if you will
need more than 10k call credits per month, etc.
5. Copy your new API key into the config.properties file
6. Update the config file with the symbols and relative excel cells to update


## Running the script
1. In Windows command line, navigate to where you saved the script, config file, and excel file.
   1. example: cd C:\Users\myuserid\crypto_stuff\
2. Execute the script: `py coinmarket.py`

## Examples
TBD

## Business Logic
1. Call API to retrieve data
2. Save data in JSON file (or just iterate through results?)
3. Open existing Excel workbook
4. Insert data into cells