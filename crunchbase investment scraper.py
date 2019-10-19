from openpyxl import Workbook
from bs4 import BeautifulSoup

def get_table(soup):
    table = soup.find("table", attrs = {'class':'card-grid'})
    return table

def get_rows(table):
    rows = soup.find_all("tr")
    return rows
    
def get_date(data):
    date = data[0].text
    return date

def get_organization_name(data):
    organization_name = data[1].text
    return organization_name

def get_series_round(data):
    text = data[3].text
    series_round = text.split('-')[0]
    return series_round

def get_dollar_amount(data):
    dollar_amount = data[4].text
    return dollar_amount

def get_investment_information(row):
    data = row.find_all('td')
    
    date = get_date(data)
    organization_name = get_organization_name(data)
    series_round = get_series_round(data)
    dollar_amount = get_dollar_amount(data)

    investment_information = [date, organization_name, series_round, dollar_amount]
    return investment_information

def create_spreadsheet():
    wb = Workbook()
    ws = wb.active
    ws.append(['Date','Organization Name','Series Round','Dollar Amount'])
    return wb

def add_to_spreadsheet(wb, investment_information):
    ws = wb.active
    ws.append(investment_information)

def save_spreadsheet(wb):
    wb.save('ForgePoint Capital.xlsx')

    
if __name__ == '__main__':
    file_names = []
    wb = create_spreadsheet()  
    for file in file_names:
        soup = BeautifulSoup(open(file),'html.parser')
        table = get_table(soup)
        rows = get_rows(table)
        for row in rows[1:]:
            investment_information = get_investment_information(row)
            add_to_spreadsheet(wb, investment_information)
    save_spreadsheet(wb)
        