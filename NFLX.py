import yfinance as yf
import xlwings as xw
from bs4 import BeautifulSoup as bs
import requests
import time
import datetime


class OptionChain:
    def __init__(self, ticker):
        self.ticker = ticker
        self.stock_info = yf.Ticker(self.ticker)
        self.expiry_date = yf.Ticker(self.ticker).options[0]
        self.option = yf.Ticker(self.ticker).option_chain(self.expiry_date)
        self.calls = self.option.calls
        self.puts = self.option.puts
        self.history = self.stock_info.history(period="1d")
        self.price = self.history["Close"].iloc[-1]
        self.selected_columns = ['strike', 'lastPrice', 'change', 'volume', 'openInterest', 'impliedVolatility']

    def get_calls(self):
        calls_selected = self.calls.loc[:, self.selected_columns]
        return calls_selected

    def get_puts(self):
        puts_selected = self.puts.loc[:, self.selected_columns]
        return puts_selected

    def get_value_from_cell(self, filename, sheetname, cell):
        wb = xw.Book(filename)
        sheet = wb.sheets[sheetname]
        value = sheet.range(cell).value
        wb.close()
        return value
 


def create_excel_file(file_name, ticker):
    try:
        wb = xw.Book(file_name)
    except FileNotFoundError:
        wb = xw.Book()

    if 'Bonds data' not in wb.sheet_names:
        wb.sheets.add('Bonds data')

    if ticker not in wb.sheet_names:
        wb.sheets.add(ticker)

    wb.save(file_name)
    return wb


def get_bonds_data(url):
    request = requests.get(url).text
    bonds = bs(request, 'html.parser')
    table = bonds.find('table', {'class': 'W(100%)'})
    rows = table.find_all('tr')
    data = []
    for row in rows:
        cols = row.find_all('td')
        row_data = [col.text for col in cols]
        data.append(row_data)
    return data


def update_bonds_sheet(bonds_sheet, data):
    for i, row in enumerate(data):
        bonds_sheet.range(f'A{1 + i}', f'H{1 + i}').value = row


def update_option_sheet(option_sheet, calls, puts, price, expiry_date):
    option_sheet.range('A5', 'G200').value = calls
    option_sheet.range('G3').value = price
    option_sheet.range('H1').value = expiry_date
    option_sheet.range('H5', 'N200').value = puts


def main():
    ticker = 'NFLX'
    try:
        #
        # Create the Excel file
        file_name = f'{ticker}.xlsx'
        wb = create_excel_file(file_name, ticker)

        # Get the bonds data
        url = 'https://finance.yahoo.com/bonds'
        bonds_data = get_bonds_data(url)

        # Add the bonds data to the Bonds sheet
        bonds_sheet = wb.sheets('Bonds data')
        update_bonds_sheet(bonds_sheet, bonds_data)

        # Create the option chain
        while True:
            try:
                initial_time = time.time()
                option_chain = OptionChain(ticker)
                expiry = OptionChain(ticker).expiry_date
                # Get the calls and puts data
                calls = option_chain.get_calls()
                puts = option_chain.get_puts()
                # Add the calls and puts data to the Option sheet
                option_sheet = wb.sheets(ticker)
                update_option_sheet(option_sheet, calls, puts, option_chain.price, expiry)
                final_time = time.time()
                total = final_time - initial_time
                Ctime = datetime.datetime.now().time()
                print(round(total, 3))
                print("Current time: {:%H:%M:%S}".format(Ctime))
                print('')
            except Exception as e:
                print(e)
    except Exception as e:
        print(e)


if __name__ == '__main__':
    main()


