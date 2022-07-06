import openpyxl
import os
import pandas as pd
from io import StringIO
import re

pattern = ['\(ADEC.*']
df = pd.read_excel(r'C:\Users\jewel.espiritu\OneDrive - AMDATEX LAS PINAS SERVICES INC\Python\py\data\processed\MergedTicket.xlsx')
df['Service subcategory->Service name'] = df['Service subcategory->Service name'].str.replace(r'\(ADEC.*', '', regex=True)
df.to_excel(r'C:\Users\jewel.espiritu\OneDrive - AMDATEX LAS PINAS SERVICES INC\Python\py\data\processed\MergedTicket.xlsx', index = False, header = True)