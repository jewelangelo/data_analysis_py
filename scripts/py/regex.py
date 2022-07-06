import re
import pandas as pd 
df = pd.read_excel(r'C:\Users\jewel.espiritu\OneDrive - AMDATEX LAS PINAS SERVICES INC\Python\py\data\processed\MergedTicket.xlsx')
regexList = ['[Ll]at[Ll]', 
             '[Dd]elta.*\Sync', 
             'FSE.*\BC', 
             '[Rr][Dd][Ss].Deployment.*[Gg][Ii][Ss]', 
             '.*[Pp][Mm][Tt].*', '[Vv][Aa].*[Ss][Cc]', 
             '.*LAT/LONG.*', 
             '[Gg][Pp][Aa][Pp].*[Ss][Cc][Aa][Nn].*',  
             '.*[Mm]icrosoft [Rr][Dd][Pp] RCE.*[Bb][Ll][Uu][Ee][Kk][Ee][Ee].*',
             '[Qq][0-9].202[0-9].[Pp][Mm].*',
             '[Rr][Tt][Mm].[Cc].*']
for i in regexList:
    print(i)
    df=df[df['Title'].str.match(i)==False]
    df.to_excel(r'C:\Users\jewel.espiritu\OneDrive - AMDATEX LAS PINAS SERVICES INC\Python\py\data\processed\MergedTicket.xlsx', index = False, header = True)
    df

subList = ['BILLS PAYMENT - RFP', 'BILLS PAYMENT - RF', 'BILLS PAYMENT - RFR']
for o in subList:
    print(o)
    df=df[df['Service subcategory->Name'].str.match(o)==False]
    df.to_excel(r'C:\Users\jewel.espiritu\OneDrive - AMDATEX LAS PINAS SERVICES INC\Python\py\data\processed\MergedTicket.xlsx', index = False, header = True)
    df   
    
service = ['Projects', 'Server and Workstation Vulnerability fixing']
for p in service:
    print(p)
    df=df[df['Service subcategory->Service name'].str.match(p)==False]
    df.to_excel(r'C:\Users\jewel.espiritu\OneDrive - AMDATEX LAS PINAS SERVICES INC\Python\py\data\processed\MergedTicket.xlsx', index = False, header = True)
    df

team = ['[Ee][Nn][Vv].*',
        '[Bb][Ss][Dd].*',
        'Information Security',]
for q in team:
    print(q)
    df=df[df['Team->Name'].str.match(q)==False]
    df.to_excel(r'C:\Users\jewel.espiritu\OneDrive - AMDATEX LAS PINAS SERVICES INC\Python\py\data\processed\MergedTicket.xlsx', index = False, header = True)
    df  
    