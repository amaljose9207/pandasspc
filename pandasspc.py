import pandas as pd
import numpy as np
import csv
import os

data = pd.read_excel (r'/Users/I347708/Desktop/html,css,js/Exportoftransactionaldata__EN.xlsx', header=4)

data.sort_values(["Process Name", "Status"], axis=0, ascending=True, inplace=True)
df = pd.DataFrame(data, columns= ['Process Name','Service Request','Start Time','Status','SID','Service Request Customer Name','Business Tenant Type','Service Request Customer ID'])

wfalms = pd.DataFrame(data, columns=['Process Name','Service Request','Start Time','Status','SID','Service Request Customer Name','Business Tenant Type','Service Request Customer ID'])
search_condition = ['BIZX GREENFIELD: IAS AND IPS INTEGRATION','SAP PROVISIONING FIND AND TRIGGER PROVISIONING PROCEDURE','SAP PROVISIONING SFSF BIZX STOCK TO CUSTOMER','SAP PROVISIONING SFSF ONB CUSTOMER TENANT SETUP','SAP PROVISIONING SFSF WFA','LMS PROVISIONING','SAP PROVISIONING SFSF JAM-BETA','SAP MULTIPOSTING TENANT PROVISIONING PROCESS','SAP SUCCESSFACTORS RECRUITING POSTING PROVISIONING SETUP','SAP PROVISIONING SFSF J2W','SAP PROVISIONING SFSF JAM WITH SCI AND IPS INTEGRATION','SAP PROVISIONING SFSF PROCESS FOR NS2',]
wfalms = wfalms[wfalms['Process Name'].str.contains('|'.join(search_condition))]

search_condition = ['LMS: INTEGRATE WITH IAS AND IPS','LIT','DE-ACTIVATE','CHANGE','RECONFIGURE','LIVE','SFSF INTEGRATE BIZX WITH IAS AND IPS','DECOMISSIONING','ASSERTION','STFK','CLEANUP','REACTIVATE','DEACTIVATE','SUBSCRIPTION','DECOMMISSIONING','HEC','DELETE','INVALIDATE','UPDATE','DECOMMISIION','REFRESH','INVALIDATION','SAC','HSHI','TERMINATE']
# search_condition = ['UPSELL','MANAGER','CLEANUP','ASSIGN ','RENEWAL ','PASSWORD','REACTIVATE','DEACTIVATE','UNASIGN','SUBSCRIPTION','DECOMMISSIONING','HEC','DELETE','INVALIDATE','UPDATE','DECOMMISIION','SFTP','REFRESH','INVALIDATION','IAS ','SAC','HSHI']
df = df[~df['Process Name'].str.contains('|'.join(search_condition))]

thelist = df.pivot_table(index=['Process Name'], aggfunc='size')
custlist = df['Service Request Customer Name'].value_counts()

print(custlist)
print(thelist)
print(df)
print(wfalms)

thelist.to_csv('pivotexport.csv')
df.to_csv('allsfrenw.csv')
wfalms.to_csv('prodprov.csv')
custlist.to_csv('custcount.csv')
print(df.to_html())

dfhtml1 = pd.read_csv('/Users/I347708/Desktop/html,css,js/prodprov.csv')
dfhtml2 = pd.read_csv('/Users/I347708/Desktop/html,css,js/allsfrenw.csv')
dfhtml3 = pd.read_csv('/Users/I347708/Desktop/html,css,js/pivotexport.csv')
dfhtml4 = pd.read_csv('/Users/I347708/Desktop/html,css,js/custcount.csv')

dfhtml1.to_html('productprovisioning.html')
dfhtml2.to_html('allsfre.html')
dfhtml3.to_html('pivot.html')
dfhtml4.to_html('custcount.html')

os.remove('Exportoftransactionaldata__EN.xlsx')
os.remove('pivotexport.csv')
os.remove('allsfrenw.csv')
os.remove('prodprov.csv')
os.remove('custcount.csv')



