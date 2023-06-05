import pandas as pd
from datetime import datetime
kpi_workbook = pd.read_excel(r'D:\Maintenance\Bloc 6\PROJECT KPI\test.xlsx', sheet_name="Sheet1")

for cell in kpi_workbook["date"]:
    data_type = type(cell)
    if data_type == str:
       date_time_object= datetime.strptime(cell , '%d/%m/%Y %H:%M:%S')
       print(date_time_object)
    else:
        print(cell)
  
