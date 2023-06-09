from datetime import datetime
import pandas as pd
from datetime import date

todays_date = pd.to_datetime(date.today())

kpi_workbook=pd.read_excel(r'D:\Maintenance\Bloc 6\PROJECT KPI\COPIL\Extraction_DI.xlsx')
kpi_workbook=pd.read_excel(r'D:\Maintenance\Bloc 6\PROJECT KPI\COPIL\Extraction_OT.xlsx')

kpi_workbook['Date création DI'] = pd.to_datetime(kpi_workbook['Date création DI'])
kpi_workbook['Anné'] = kpi_workbook['Date création DI'].dt.year
kpi_workbook['Moi'] = kpi_workbook['Date création DI'].dt.month
kpi_workbook['Jour'] = kpi_workbook['Date création DI'].dt.day
backlog_age= todays_date - kpi_workbook['Date création DI']

for index, row in kpi_workbook.iterrows():
    if backlog_age[index] < pd.Timedelta('90 days'):
       kpi_workbook.at[index, 'Backlog Age'] = 'Backlog < 3 Mois'
       print('Backlog < 3 Mois')
    elif backlog_age[index] > pd.Timedelta('180 days'):
         kpi_workbook.at[index, 'Backlog Age'] = 'Backlog > 6 Mois'
         print('Backlog > 6 Mois') 
    else:
       backlog_age[index] > pd.Timedelta('30 days') and backlog_age[index] < pd.Timedelta('180 days')
       kpi_workbook.at[index, 'Backlog Age'] = '3 Mois < Backlog < 6 Mois'
       print('3 Mois < Backlog < 6 Mois')

selected_columns = ['Date création DI', 'Anné', 'Moi', 'Jour', 'Backlog Age']
kpi_workbook[selected_columns].to_excel(r'D:\Maintenance\Bloc 6\PROJECT KPI\COPIL\fin.xlsx')



