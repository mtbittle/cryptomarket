#This was tested with Python 3.6

from requests import Request, Session
from requests.exceptions import ConnectionError, Timeout, TooManyRedirects
import json
import configparser
import collections
from openpyxl import load_workbook


# function to load the config file so we can have user-specific API, crypto currencies, excel fields
def readconfig():
    print("Reading configuration file to set up user-specific values.")
    config = configparser.ConfigParser()
    config.read('config.properties')
    return config


# API call to CoinMarket - assumes config.properties has all props filled in
def call_api(crypto_symbols, base_url, endpoint, key):
    url = base_url + endpoint
    parameters = {
        'symbol': crypto_symbols
    }
    headers = {
        'Accepts': 'application/json',
        'X-CMC_PRO_API_KEY': key,
    }

    session = Session()
    session.headers.update(headers)

    try:
        response = session.get(url, params=parameters)
        json_response = json.loads(response.text)
        print(json.dumps(json_response, indent=4, sort_keys=True))
        return json_response
    except (ConnectionError, Timeout, TooManyRedirects) as e:
        print(e)


# function to create a dict for each symbol and the field in excel to update
def get_excel_mapping(symbols_list, excel_list):
    symbol_map = {symbols_list[i]: excel_list[i] for i in range(len(symbols_list))}
    print("Excel field data: " + str(symbol_map))
    return symbol_map


# funtion to pull out the price data for each crypto returned from the API call
def parse_json_data(data):
    keys = data['data'].keys()
    crypto_dict = {}
    for key in keys:
        crypto_dict[key] = data['data'][key]['quote']['USD']['price']

    print('crypto price data: ' + str(crypto_dict))
    return crypto_dict


# function to combine the price data and the excel fields into one dict before updating the excel sheet
def get_combined_data(symbols_dict, crypto_dict):
    full_crypto_excel_dict = collections.defaultdict(dict)
    for key in symbols_dict.keys():
        full_crypto_excel_dict[key]['price'] = crypto_dict[key]
        full_crypto_excel_dict[key]['field'] = symbols_dict[key]

    return full_crypto_excel_dict


# function to update the existing excel sheet with new data
def update_excel_file(excel_file, full_data):
    wb = load_workbook(excel_file)
    sheet = wb.worksheets[0]
    for key in combined_data.keys():
        print('field: ' + combined_data[key]['field'])
        column = sheet[combined_data[key]['field']].value = combined_data[key]['price']
    wb.save(excel_file)


# MAIN SCRIPT FLOW
if __name__ == '__main__':
    # load config.properties file for user-specific props
    configurer = readconfig()
    use_test_data = configurer.get('CoinMarketSection', 'api.use.local.test.data')
    symbols = configurer.get('CryptoInfoSection', 'crypto.symbols')
    api_base_url = configurer.get('CoinMarketSection', 'api.base.url')
    api_endpoint = configurer.get('CoinMarketSection', 'api.endpoint')
    api_key = configurer.get('CoinMarketSection', 'api.key')
    # IF CONFIG FILE HAS 'true' FOR property 'api.use.local.test.data', THEN NO CALL TO REAL API WILL BE USED
    if use_test_data == 'true':
        f = open('testquotes.json')
        json_data = json.load(f)
        print('loaded from test data file')
    else:
        # call api to get crypto data in json format - see testquotes.json
        json_data = call_api(symbols, api_base_url, api_endpoint, api_key)
        print('called API')

    # parse the json data into a dict so we have info for each crypto
    crypto_prices = parse_json_data(json_data)
    print(crypto_prices)
    # print(json_data['data'].keys())
    # split out the comma-separated Strings into lists
    symbol_list = symbols.split(",")
    excel_fields = configurer.get('ExcelSection', 'excel.field.map')
    excel_fields_list = excel_fields.split(",")
    symbols_map = get_excel_mapping(symbol_list, excel_fields_list)

    combined_data = get_combined_data(symbols_map, crypto_prices)
    # print(crypto_prices)
    print('using data for Excel: ' + str(combined_data))

    excel_file_name = configurer.get('ExcelSection', 'excel.file.name')
    print('updating excel file...')
    update_excel_file(excel_file_name, combined_data)
