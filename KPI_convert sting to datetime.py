import pandas as pd
from datetime import datetime
kpi_workbook = pd.read_excel(r'D:\Maintenance\Bloc 6\PROJECT KPI\COPIL\test.xlsx', sheet_name="Sheet1")
i=1

for c in kpi_workbook["date"]:
    date_time_object= datetime.strptime(c , '%d/%m/%Y %H:%M:%S')
    data_type = type(c)
    if data_type == str:
       #date_time_object= datetime.strptime(c , '%d/%m/%Y %H:%M:%S')
       kpi_workbook.cell(row=1+i, column=2).value = date_time_object.strftime('%x %X')
       kpi_workbook.save(r'D:\Maintenance\Bloc 6\PROJECT KPI\COPIL\test.xlsx', sheet_name="Sheet1")
       i=i+1
       print(date_time_object)
    else:
        kpi_workbook.appendcell(row=1+i, column=2).value = date_time_object.strftime('%x %X')
        kpi_workbook.save(r'D:\Maintenance\Bloc 6\PROJECT KPI\COPIL\test.xlsx', sheet_name="Sheet1")
        i=i+1
        print(c)
  