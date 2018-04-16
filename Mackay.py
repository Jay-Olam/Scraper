#! python3.6

import bs4 as bs
import urllib
import urllib.request
import pandas as pd
from bs4 import BeautifulSoup
import datetime
import openpyxl as op
import xlrd
import xlwt
import requests
from openpyxl.styles import Font

def full_scrape(filepath):

    def website_scrape(x):
        url = x
        html = requests.get(url).content
        df_list = pd.read_html(html)
        df_list=df_list[0]
        return (df_list)
        
    Mackay=website_scrape('https://nqbp.com.au/operations/shipping-schedules')

    def sugar_filter(df):
        contains='Sugar|sugar'
        df=df[df['CARGO'].str.contains(contains)]
        return(df)

    Mackay=sugar_filter(Mackay)

    Mackay=Mackay.fillna('No info')

    def add_time(df):  
        
        now=datetime.datetime.now()
        current_time=now.strftime("%Y-%m-%d %H:%M:%S")
        time_col= [current_time]*(len(df))
        df['Time']=time_col

        return(df)

    Mackay=add_time(Mackay)

    def sheet_reader(file_name,sheet_name):
        df=pd.read_excel(file_name, sheet_name=sheet_name, skiprows=1)
        return(df)

    #Loading in historical data
    Overall=sheet_reader(filepath, 'Overall')

    #Loading in yesterday data
    Yesterday=sheet_reader(filepath, 'Today')

    Combined=Overall.append(Yesterday)

    #Sending to Excel file
    writer=pd.ExcelWriter(filepath)

    #Sending to main history sheet
    Combined.to_excel(writer, 'Overall', startrow=1, startcol=0, index=False)

    #Sending to today sheet
    Mackay.to_excel(writer, 'Today', startrow=1, startcol=0, index=False)

    writer.save()

    def main_sheet_formatter(sheet):

        sheet['A1']='Mackay' 
        sheet['A1'].font=Font(size=16)
        
        column_dim_1={'A': 12, 'B': 8, 'C':18, 'D':18, 'E':12, 'F':20, 'G':12, 'H':12 , 'I':12, 'J': 12, 'K': 20 }
       

        for key in column_dim_1:
            sheet.column_dimensions[key].width=column_dim_1[key]

            now=datetime.datetime.now()
            current_time=now.strftime("%Y-%m-%d %H:%M:%S")

            sheet['K1']='{}'.format(current_time)

    def doc_formatter(x):

        #Formatting the workbook
        wb=op.load_workbook(filename=x)

        Overall=wb.worksheets[0]
        Today=wb.worksheets[1]

        main_sheet_formatter(Overall)
        main_sheet_formatter(Today)

        wb.save(x)
        
    doc_formatter(filepath)

    print('Done')

full_scrape(r'S:\Port tracking\Mackay-Aus.xlsx')


