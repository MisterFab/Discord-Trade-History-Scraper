import os
import pandas as pd
import requests
from datetime import datetime, timedelta
import logging

API_TOKEN = os.environ.get('POLYGON_TOKEN')
MAX_RETRIES = 10
API_BASE_URL = 'https://api.polygon.io/v2/aggs/ticker'

logging.basicConfig(level=logging.INFO)

def convert_strike(strike):
    s = format(strike, '08.2f')
    return s.replace('.', '') + '0'

def create_url(symbol, expiry, type, strike, start_date, token):

    while expiry < datetime.strptime(start_date, '%Y-%m-%d'):
        expiry = expiry.replace(year=expiry.year + 1)

    expiry_for_ticker = datetime.strftime(expiry, '%y%m%d')
    expiry_for_end_date = datetime.strftime(expiry, '%Y-%m-%d')
    strike_converted = convert_strike(strike)
    ticker = f"{symbol}{expiry_for_ticker}{type}{strike_converted}"
    url = f"{API_BASE_URL}/O:{ticker}/range/1/minute/{start_date}/{expiry_for_end_date}?adjusted=true&sort=asc&limit=5000&apiKey={token}"
    return url

def get_data(expiry, start_date, start_time, symbol, type, strike, token):
    start_time_trade = datetime.strptime(f"{start_date}{start_time}", '%Y-%m-%d%H%M%S')
    
    url = create_url(symbol, expiry, type, strike, start_date, token)

    response = requests.get(url).json()
    count = 0

    while response.get("resultsCount", 0) == 0 and count < MAX_RETRIES:
        expiry += timedelta(days=1)
        url = create_url(symbol, expiry, type, strike, start_date, token)
        response = requests.get(url).json()
        count += 1

    try:
        high = max([i['h'] for i in response.get("results", []) if datetime.utcfromtimestamp(i['t'] / 1000.0) >= start_time_trade])
        low = min([i['l'] for i in response.get("results", []) if datetime.utcfromtimestamp(i['t'] / 1000.0) >= start_time_trade])
        entry = next((i['h'] for i in response.get("results", []) if datetime.utcfromtimestamp(i['t'] / 1000.0) >= start_time_trade), None)
    except Exception as e:
        return None, None, None, datetime.strftime(expiry, '%y%m%d'), url

    return high, low, entry, datetime.strftime(expiry, '%y%m%d'), url

def process_row(row):
    global processed_rows
    url = ''
    try:
        expiry = datetime.strptime(row['Expiry'], '%d%b%y') if pd.notna(row['Expiry']) else datetime.strptime(str(row['Date']),'%Y-%m-%d')
        high, low, entry, expiry_for_ticker, url = get_data(expiry, str(row['Date']), str(row['Time'][:6].replace(':', '')), row['Symbol'], row['Type'], float(row['Strike']), API_TOKEN)
        
        max_profit = float(high) / float(entry) - 1
        processed_rows += 1
        if processed_rows % 10 == 0:
            logging.info(f"Processed {processed_rows} rows so far.")
        return high, low, entry, expiry_for_ticker, max_profit
    except Exception as e:
        logging.error(f"Error processing url: {url}")
        return None, None, None, datetime.strftime(expiry, '%y%m%d'), None


def main():
    xls = pd.ExcelFile('trades_scraped.xlsx')
    writer = pd.ExcelWriter('trades_finished.xlsx')

    for sheet_name in xls.sheet_names:
        logging.info(f"Processing sheet: {sheet_name}.")
        contracts_file = pd.read_excel(xls, sheet_name=sheet_name)
        df = pd.DataFrame(contracts_file)
        logging.info(f"Sheet {sheet_name} has {len(df)} rows to process.")
        
        global processed_rows 
        processed_rows = 0
        result = df.apply(process_row, axis=1, result_type='expand')
        df['High'], df['Low'], df['Entry'], df['Expiry'], df['Max Profit'] = result[0], result[1], result[2], result[3], result[4]
        
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        logging.info(f"Completed processing for sheet: {sheet_name}.")

    logging.info("All sheets processed. Writing to trades_finished.xlsx.")
    writer.save()

if __name__ == '__main__':
    main()