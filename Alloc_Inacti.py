### Allocation Inactivation ###
# Initial Coding : Sunil Kumar Mishra
# Created On : 05/12/2018 
# Description: Will automate the task of Searching the Id for the Names 
#              of the product recieved in the mail, Deslecting the Activate 
#              Checkbox for the product IDs and Deactivating related allocations.

import simple_salesforce
import pandas as pd
from simple_salesforce import Salesforce
import os


os.chdir('C:\\Users\\sunil.kumar.mishra.m\\Desktop\\Python\\Regenerron Use Case\\Allocation Inactivation')

sf = Salesforce(username="support1@regeneron.com.full" , password="Regeneron2",security_token='', sandbox=True)
print('Login Successfull')

df = pd.read_excel('Prod_code.xlsx')
prod= df['Product_code'].tolist()
print(df)
Id= []
Alloc_Id=[]
Prod_Id=[]

for i in range(len(df)):
    query=sf.query("Select Id,Product_vod__c from Inventory_Order_Allocation_vod__c where Name = '%s'"% prod[i])
    for record in query['records']:
        Alloc_Id.append(record['Id'])
        Prod_Id.append(record['Product_vod__c'])
        
        for i in range(len(Alloc_Id)):
            sf.Inventory_Order_Allocation_vod__c.update(Alloc_Id[i],{'Active_vod__c':False})

   # query2=sf.query("Select Id from Product_vod__c where Id = '%s'"% Prod_Id[i])
    for i in range(len(Prod_Id)):
            sf.Product_vod__c.update(Prod_Id[i], {'Company_Product_vod__c':False})
print("Allocations along with product Inactivated")



