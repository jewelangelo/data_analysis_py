import pandas as pd
import shutil
 
path = r'C:\repo\SLA_PowerBI.xlsx'
 
writer = pd.ExcelWriter(path, engine='xlsxwriter')

paretodf = pd.read_excel(r'C:\repo\SLA_Pareto.xlsx') #<-- Excel file from SLA_Pareto.ipynb
slaperfdf = pd.read_excel(r'C:\repo\SLA_Perf.xlsx') #<-- Excel file from SLA_Perf.ipynb
trenddf = pd.read_excel(r'C:\repo\Trend.xlsx') # <-- Excel file from Trend.ipynb
beyonddf = pd.read_excel(r'C:\repo\SLA_Pareto_BeyondTAT.xlsx') #<-- Excel file from SLA_Pareto_BeyondTAT.ipynb
servicedf = pd.read_excel(r'C:\repo\SLA_Service.xlsx') #<-- Excel file from SLA_ServiceCount.ipynb
filingdf = pd.read_excel(r'C:\repo\SLA_Filing.xlsx') #<-- Excel file from SLA_ServiceCount.ipynb
 
paretodf.to_excel(writer, sheet_name = 'Pareto')
slaperfdf.to_excel(writer, sheet_name = 'SLA')
trenddf.to_excel(writer, sheet_name = 'TREND')
beyonddf.to_excel(writer, sheet_name = 'Beyond')
servicedf.to_excel(writer, sheet_name= 'Service')
filingdf.to_excel(writer, sheet_name= 'Filing')
writer.save()

src = r'C:\repo\SLA_PowerBI.xlsx'
dst = r'C:\Users\jewel.espiritu\OneDrive - AMDATEX LAS PINAS SERVICES INC\SLA\repo\SLA_PowerBI.xlsx' #<-- This will be the location of Power BI data source

shutil.copyfile(src, dst)