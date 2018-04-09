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
from openpyxl.styles import Font

def sheet_reader(file_name,sheet_name):
    df=pd.read_excel(file_name, sheet_name=sheet_name)
    return(df)

def initial_scrape_expected(x):
    #Accessing the website and retrieve the table

    url=x
    html=urllib.request.urlopen(url).read()
    soup=bs.BeautifulSoup(html, 'html.parser')
    table=soup.findAll('tr')

    #Placing data in a pandas data frame

    shipdata=[]
    for record in table:
        shipdata.append(record.text)

    rows=[]
    for row in shipdata:
        rows.append(row.splitlines())

    ship_data=pd.DataFrame(rows)

    #Formatting data frame

    header=ship_data.iloc[0]
    new_header=[]

    for title in header:
        new_header.append(title.strip())

    ship_data=ship_data[1:]
    ship_data.columns=new_header
    
    sugar=[ship_data['GOODS'].str.contains("SUCRE")]
    
    ship_data=ship_data[ship_data['GOODS'].str.contains("SUCRE")]

    #Adding a time column so we can track dates
    now=datetime.datetime.now()
    current_time=now.strftime("%Y-%m-%d %H:%M:%S")

    time_col= [current_time]*(len(ship_data))

    ship_data['Time']=time_col

    ship_data=ship_data.drop('', axis=1)
    
    return(ship_data)

def initial_scrape_offshore(x):
    #Accessing the website and retrieve the table

    url=x
    html=urllib.request.urlopen(url).read()
    soup=bs.BeautifulSoup(html, 'html.parser')
    table=soup.findAll('tr')

    #Placing data in a pandas data frame

    shipdata=[]
    for record in table:
        shipdata.append(record.text)

    rows=[]
    for row in shipdata:
        rows.append(row.splitlines())

    ship_data=pd.DataFrame(rows)

    header=[]
    for col_title in range(1,9):
        header.append(ship_data.iloc[0][col_title])

    cols = [0,1,3]
    ship_data.drop(ship_data.columns[cols],axis=1,inplace=True)
    ship_data.drop([0], axis=0, inplace=True)

    ship_data.columns=header
    
    new_header=[]
    
    for title in header:
        new_header.append(title.strip())

    ship_data.columns=new_header

    sugar=[ship_data['GOODS'].str.contains("SUCRE")]

    ship_data=ship_data[ship_data['GOODS'].str.contains("SUCRE")]

    #Adding a time column so we can track dates
    now=datetime.datetime.now()
    current_time=now.strftime("%Y-%m-%d %H:%M:%S")

    time_col= [current_time]*(len(ship_data))

    ship_data['Time']=time_col

    return(ship_data)

def initial_scrape_indock(x):
    
    url=x
    html=urllib.request.urlopen(url).read()
    soup=bs.BeautifulSoup(html, 'html.parser')
    table=soup.findAll('tr')

    #Placing data in a pandas data frame

    shipdata=[]
    for record in table:
        shipdata.append(record.text)

    rows=[]
    for row in shipdata:
        rows.append(row.splitlines())

    ship_data=pd.DataFrame(rows)

    header=[]
    for col_title in range(1,10):
        header.append(ship_data.iloc[0][col_title])

    cols = [0,2,4]
    ship_data.drop(ship_data.columns[cols],axis=1,inplace=True)
    ship_data.drop([0], axis=0, inplace=True)

    ship_data.columns=header

    new_header=[]

    for title in header:
        new_header.append(title.strip())

    ship_data.columns=new_header

    sugar=[ship_data['GOODS'].str.contains("SUCRE")]

    ship_data=ship_data[ship_data['GOODS'].str.contains("SUCRE")]

    #Adding a time column so we can track dates
    now=datetime.datetime.now()
    current_time=now.strftime("%Y-%m-%d %H:%M:%S")

    time_col= [current_time]*(len(ship_data))

    ship_data['Time']=time_col

    return(ship_data)

def main_sheet_formatter(sheet):

    sheet['A1']='Expected' 
    sheet['A1'].font=Font(size=16)
    
    sheet['L1']='Offshore'
    sheet['L1'].font=Font(size=16)
    
    sheet['V1']='Indock'
    sheet['V1'].font=Font(size=16)

    sheet['H1']='Last Updated'

    now=datetime.datetime.now()
    current_time=now.strftime("%Y-%m-%d %H:%M:%S")

    sheet['I1']='{}'.format(current_time)

    column_dim_1={'A': 20, 'B': 18, 'C':12, 'D':10, 'E':15, 'F':15, 'G': 18, 'H':20, 'I':18, 'J':20}
    column_dim_2={'L': 20, 'M': 18, 'N':12, 'O':10, 'P':15, 'Q':15, 'R': 15, 'S':18, 'T':20}
    column_dim_3={'V': 15, 'W': 20, 'X':10, 'Y':18, 'Z':15, 'AA':15, 'AB': 15, 'AC':18, 'AD':15, 'AE':20}

    for key in column_dim_1:
        sheet.column_dimensions[key].width=column_dim_1[key]
        
    for key in column_dim_2:
        sheet.column_dimensions[key].width=column_dim_2[key]
        
    for key in column_dim_3:
        sheet.column_dimensions[key].width=column_dim_3[key]

def doc_formatter(x):

    #Formatting the workbook
    wb=op.load_workbook(filename=x)

    Overall=wb.worksheets[0]
    Today=wb.worksheets[1]

    main_sheet_formatter(Overall)
    main_sheet_formatter(Today)

    wb.save(x)

def full_scrape(filepath):
    #Initital scraping todays data
    Expected=initial_scrape_expected('https://www.portdebejaia.dz/index.php/en/vessles-s-position/expected')
    Offshore=initial_scrape_offshore('https://www.portdebejaia.dz/index.php/en/vessles-s-position/offshore')
    Indock=initial_scrape_indock('https://www.portdebejaia.dz/index.php/en/vessles-s-position/in-dock')

    #Loading in historical data
    Over_Ex=sheet_reader(filepath, 'Exp_His')
    Over_Os=sheet_reader(filepath, 'Offs_His')
    Over_Id=sheet_reader(filepath, 'Indock_His')

    #Loading in yesterday's data
    Yes_Ex=sheet_reader(filepath, 'Exp_Today')
    Yes_Os=sheet_reader(filepath, 'Offs_Today')
    Yes_Id=sheet_reader(filepath, 'Indock_Today')

    #Combining the dataframes
    Comb_Ex=Over_Ex.append(Yes_Ex)
    Comb_Os=Over_Os.append(Yes_Os)
    Comb_Id=Over_Id.append(Yes_Id)

    #Sending to Excel file
    writer=pd.ExcelWriter(filepath)

    #Sending to main history sheet
    Comb_Ex.to_excel(writer, 'Overall', startrow=1, startcol=0, index=False)
    Comb_Os.to_excel(writer, 'Overall', startrow=1, startcol=len(Expected.columns) + 1, index=False)
    Comb_Id.to_excel(writer, 'Overall', startrow=1, startcol=len(Expected.columns)+len(Offshore.columns) +2, index=False)

    #Sending to main today sheet
    Expected.to_excel(writer, 'Today', startrow=1, startcol=0, index=False)
    Offshore.to_excel(writer, 'Today', startrow=1, startcol=len(Expected.columns) + 1, index=False)
    Indock.to_excel(writer, 'Today', startrow=1, startcol=len(Expected.columns)+len(Offshore.columns) +2, index=False)

    #Sending to subsheets
    Comb_Ex.to_excel(writer, 'Exp_His', startrow=1, startcol=0, index=False)
    Expected.to_excel(writer, 'Exp_Today', startrow=1, startcol=0, index=False)
    Comb_Os.to_excel(writer, 'Offs_His', startrow=1, startcol=0, index=False)
    Offshore.to_excel(writer, 'Offs_Today', startrow=1, startcol=0, index=False)
    Comb_Id.to_excel(writer, 'Indock_His', startrow=1, startcol=0, index=False)
    Indock.to_excel(writer, 'Indock_Today', startrow=1, startcol=0, index=False)

    writer.save()

    doc_formatter(filepath)
    
#full_scrape(r'C:\Users\jay.haran\Documents\PC_scrape\Debajia.xlsx')

full_scrape(r'S:\Port tracking\Debajia.xlsx')
