import googlefinance.client as gf
import pyexcel as pe
import numpy as np
import pandas as pd
import os

def get_stock_price(istock_name):

    # Price from google finance, last 3 months average
    param = {
        'q': istock_name, # Stock symbol (ex: "AAPL")
        'i': "86400", # Interval size in seconds ("86400" = 1 day intervals)
        'x': "HEL", # Stock exchange symbol on which stock is traded (ex: "NASD")
        'p': "1Y" # Period (Ex: "1Y" = 1 year)
    }
    df = gf.get_price_data(param)
    stock_price = (df.Close.tolist())
    averaged_price = np.average(stock_price)

    return averaged_price

def change_working_directory(iindustry):
    print(os.getcwd())
    os.chdir(os.path.join(os.path.abspath(os.path.curdir), 'companies\\' + iindustry))

def list_industries():
    industries_list = []
    for name in os.listdir(os.chdir(os.path.join(os.path.abspath(os.path.curdir), u'companies'))):
        industries_list.append(name)
    return industries_list

def list_all_companies():
    name_sheet = pe.get_sheet(file_name="company_names.ods", start_row=1)
    name_list = list(name_sheet.column_at(0))
    return name_list

def get_stock_data(istock_name, iindustry):

    stock_sheet = pe.get_sheet(file_name=istock_name + ".ods")
    stock_price = get_stock_price(istock_name)
    stock_eps = stock_sheet['B6']
    stock_shares = stock_sheet['B7']
    return stock_price, stock_eps, stock_shares

def swap_to_stock_name(istock_name):

    i = 0
    name_sheet = pe.get_sheet(file_name="company_names.ods", start_row=1)
    name_list = list_all_companies()
    stock_list = list(name_sheet.column_at(1))
    for name in name_list:
        if istock_name == name_list[i]:
            return stock_list[i]
        else:
            i = i + 1



industry = "Transport"
stock_name = swap_to_stock_name("Finnair Oyj")
work_dir=os.getcwd()
print(work_dir)
change_working_directory(industry)
print(os.getcwd())
stock_price, stock_eps, stock_shares = get_stock_data(stock_name, industry)
print(stock_name)
print(stock_price)
print(stock_eps)
print(stock_shares)
# print(stock_name+" P/E: ",stock_price/stock_eps)

