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

    Operating=website_scrape('http://www.tartousport.gov.sy/ship_mov.php')
    Hyphen=website_scrape('http://www.tartousport.gov.sy/ship_waiting.php')

    from googletrans import Translator

    translator=Translator()

    #Translating

    def trans(text):
        english=translator.translate(text).text

        return(english)

    for column in range (0,len(Operating.columns)):
            Operating[column]=Operating[column].apply(trans)

    for column in range (0,len(Hyphen.columns)):
            Hyphen[column]=Hyphen[column].apply(trans)


    def header_adjust(df):
        new_header=df.iloc[0]
        df=df[1:]
        df.columns=new_header

        return(df)

    Operating=header_adjust(Operating)
    Hyphen=header_adjust(Hyphen)

    def sugar_filter(df):
        contains="Sugar|sugar|SUGAR|oils"
        df=df[df['Type of goods'].str.contains(contains)]
        return(df)

    Operating=sugar_filter(Operating)
    Hyphen=sugar_filter(Hyphen)

    def add_time(df):  
        
        now=datetime.datetime.now()
        current_time=now.strftime("%Y-%m-%d %H:%M:%S")
        time_col= [current_time]*(len(df))
        df['Time']=time_col

        return(df)

    Operating=add_time(Operating)
    Hyphen=add_time(Operating)

    #Read in existing spreadsheet
    def sheet_reader(file_name,sheet_name):
        df=pd.read_excel(file_name, sheet_name=sheet_name)
        return(df)

    Oper_His=sheet_reader(filepath, 'Oper_His')
    Oper_Tod=sheet_reader(filepath, 'Oper_Tod')
    Hyphen_His=sheet_reader(filepath, 'Hyphen_His')
    Hyphen_Tod=sheet_reader(filepath, 'Hyphen_Tod')

    Comb_Op=Oper_His.append(Oper_Tod)
    Comb_Hyphen=Hyphen_His.append(Hyphen_Tod)

    writer=pd.ExcelWriter(filepath)

    #Main excel sheets
    Comb_Op.to_excel(writer, 'Overall', startrow=1, startcol=0, index=False)
    Comb_Hyphen.to_excel(writer, 'Overall', startrow=1, startcol=len(Operating.columns) + 1, index=False)

    Operating.to_excel(writer, 'Today', startrow=1, startcol=0, index=False)
    Hyphen.to_excel(writer, 'Today', startrow=1, startcol=len(Operating.columns) + 1, index=False)

    #Sending to subsheets
    Comb_Op.to_excel(writer, 'Oper_His', startrow=1, startcol=0, index=False)
    Operating.to_excel(writer, 'Oper_Tod', startrow=1, startcol=0, index=False)
    Comb_Hyphen.to_excel(writer, 'Hyphen_His', startrow=1, startcol=0, index=False)
    Hyphen.to_excel(writer, 'Hyphen_Tod', startrow=1, startcol=0, index=False)

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

        sheet['A1']='Operating Vessels' 
        sheet['A1'].font=Font(size=16)

        sheet['H1']='Last Updated'

        now=datetime.datetime.now()
        current_time=now.strftime("%Y-%m-%d %H:%M:%S")

        sheet['I1']='{}'.format(current_time)
        
        sheet['K1']='Ship Hyphenation'
        sheet['K1'].font=Font(size=16)

        column_dim_1={'A': 18, 'B': 18, 'C':18, 'D':18, 'E':18, 'F':18, 'G':18, 'H':18, 'I':18}
        column_dim_2={'K': 18, 'L': 18, 'M':18, 'N':18, 'O':18, 'P':18, 'Q':18, 'R':18, 'S':18}

        for key in column_dim_1:
            sheet.column_dimensions[key].width=column_dim_1[key]
            
        for key in column_dim_2:
            sheet.column_dimensions[key].width=column_dim_2[key]
            
            
    doc_formatter(filepath)

#full_scrape(r'C:\Users\jay.haran\Documents\PC_scrape\Tartoum.xlsx')

full_scrape(r'S:\Port tracking\Tartoum.xlsx')













