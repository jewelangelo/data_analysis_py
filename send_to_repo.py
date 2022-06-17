import pandas as pd
import shutil
 
path = r'C:\repo\SLA_PowerBI.xlsx'
 
writer = pd.ExcelWriter(path, engine='xlsxwriter')

paretodf = pd.read_excel(r'C:\repo\SLA_Pareto.xlsx')
slaperfdf = pd.read_excel(r'C:\repo\SLA_Perf.xlsx')
trenddf = pd.read_excel(r'C:\repo\Trend.xlsx')
 
paretodf.to_excel(writer, sheet_name = 'Pareto')
slaperfdf.to_excel(writer, sheet_name = 'SLA')
trenddf.to_excel(writer, sheet_name = 'TREND')
writer.save()

src = r'C:\repo\SLA_PowerBI.xlsx'
dst = r'C:\Users\jewel.espiritu\OneDrive - AMDATEX LAS PINAS SERVICES INC\SLA\repo\SLA_PowerBI.xlsx'

shutil.copyfile(src, dst)