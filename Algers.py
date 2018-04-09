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

    #Scrape today data
    def website_scrape(x):
        url = x
        html = requests.get(url).content
        df_list = pd.read_html(html)
        df= df_list[-1]
        df=df.fillna('-')
        return (df)

    harbour_today=website_scrape('http://www.portalger.com.dz/situation-du-port/navires-en-rade')
    dock_today=website_scrape('http://www.portalger.com.dz/situation-du-port/navires-a-quai')
    out_today=website_scrape('http://www.portalger.com.dz/situation-du-port/navires-sortis')

    #Filter for sugar
    def sugar_filter(df):
        contains='Sucre|sucre|SUCRE'
        df=df[df['Marchandise'].str.contains(contains)]
        return(df)

    harbour_today=sugar_filter(harbour_today)
    dock_today=sugar_filter(dock_today)
    out_today=sugar_filter(out_today)

    #Adding time column
    def add_time(df):  
        
        now=datetime.datetime.now()
        current_time=now.strftime("%Y-%m-%d %H:%M:%S")
        time_col= [current_time]*(len(df))
        df['Time']=time_col

        return(df)

    harbour_today=add_time(harbour_today)
    dock_today=add_time(dock_today)
    out_today=add_time(out_today)

    #Read in existing spreadsheet
    def sheet_reader(file_name,sheet_name):
        df=pd.read_excel(file_name, sheet_name=sheet_name)
        return(df)

    #Loading in historical data
    Over_Har=sheet_reader(filepath, 'Harbour_His')
    Over_Dock=sheet_reader(filepath, 'Dock_His')
    Over_Out=sheet_reader(filepath, 'Out_His')

    #Loading in yesterday's data
    Yes_Har=sheet_reader(filepath, 'Harbour_Today')
    Yes_Dock=sheet_reader(filepath, 'Dock_Today')
    Yes_Out=sheet_reader(filepath, 'Out_Today')

    #Combining dataframes
    Comb_Har=Over_Har.append(Yes_Har)
    Comb_Dock=Over_Dock.append(Yes_Dock)
    Comb_Out=Over_Out.append(Yes_Out)

    #Sending to Excel file
    writer=pd.ExcelWriter(filepath)

    #Sending to main history sheet
    Comb_Har.to_excel(writer, 'Overall', startrow=1, startcol=0, index=False)
    Comb_Dock.to_excel(writer, 'Overall', startrow=1, startcol=len(harbour_today.columns) + 1, index=False)
    Comb_Out.to_excel(writer, 'Overall', startrow=1, startcol=len(harbour_today.columns)+len(dock_today.columns) +2, index=False)

    #Sending to main today sheet
    harbour_today.to_excel(writer, 'Today', startrow=1, startcol=0, index=False)
    dock_today.to_excel(writer, 'Today', startrow=1, startcol=len(harbour_today.columns) + 1, index=False)
    out_today.to_excel(writer, 'Today', startrow=1, startcol=len(harbour_today.columns)+len(dock_today.columns) +2, index=False)

    #Sending to subsheets
    Comb_Har.to_excel(writer, 'Harbour_His', startrow=1, startcol=0, index=False)
    harbour_today.to_excel(writer, 'Harbour_Today', startrow=1, startcol=0, index=False)
    Comb_Dock.to_excel(writer, 'Dock_His', startrow=1, startcol=0, index=False)
    dock_today.to_excel(writer, 'Dock_Today', startrow=1, startcol=0, index=False)
    Comb_Out.to_excel(writer, 'Out_His', startrow=1, startcol=0, index=False)
    out_today.to_excel(writer, 'Out_Today', startrow=1, startcol=0, index=False)

    writer.save()

    #Formatting the file

    def doc_formatter(x):

        #Formatting the workbook
        wb=op.load_workbook(filename=x)

        Overall=wb.worksheets[0]
        Today=wb.worksheets[1]

        main_sheet_formatter(Overall)
        main_sheet_formatter(Today)

        wb.save(x)

    def main_sheet_formatter(sheet):

        sheet['A1']='Harbour' 
        sheet['A1'].font=Font(size=16)
        
        sheet['H1']='Dock'
        sheet['H1'].font=Font(size=16)
        
        sheet['Q1']='Out'
        sheet['Q1'].font=Font(size=16)

        sheet['E1']='Last Updated'

        now=datetime.datetime.now()
        current_time=now.strftime("%Y-%m-%d %H:%M:%S")

        sheet['F1']='{}'.format(current_time)
        

        column_dim_1={'A': 12, 'B': 15, 'C':20, 'D':18, 'E':12, 'F':20}
        column_dim_2={'H': 12, 'I': 15, 'J':20, 'K':15, 'L':15, 'M':10, 'N':12, 'O': 20}
        column_dim_3={'Q': 12, 'R': 15, 'S':20, 'T':15, 'U':15, 'V':10, 'w': 20}

        for key in column_dim_1:
            sheet.column_dimensions[key].width=column_dim_1[key]
            
        for key in column_dim_2:
            sheet.column_dimensions[key].width=column_dim_2[key]
            
        for key in column_dim_3:
            sheet.column_dimensions[key].width=column_dim_3[key]
            
    doc_formatter(filepath)

#full_scrape(r'C:\Users\jay.haran\Documents\PC_scrape\Algers.xlsx')

full_scrape(r'S:\Port tracking\Algers.xlsx')
