# -*- coding: utf-8 -*-
"""
Created on Thu Sep 23 11:55:57 2021

@author: vishnu.mohan
"""

import pandas as pd
import numpy as np
import urllib
#import xlsxwriter
import datetime
import tkinter as tk
from tkinter import filedialog
#import time
root=tk.Tk()

#Select delivery file
path=filedialog.askopenfilename(title='Choose the Delivery File')

bj_file=pd.read_excel(path,header=[1], sheet_name="TOP 200 - BJ's")
disc_file=pd.read_excel(path,header=[1], sheet_name="TOP 200 - Discount Tire")
sams_file=pd.read_excel(path,header=[1], sheet_name="TOP 200 - Sam's Club")
cvr=pd.read_excel(path,header=[1], sheet_name="Cover Page")
status=pd.read_excel(path, sheet_name="Status-By-Date")
summary=pd.read_excel(path, sheet_name="Summary")


#Select BJ's File
bj_crawl_path=filedialog.askopenfilename(title='Choose Bjs Output')
bj_crawl=pd.read_csv(bj_crawl_path, sep = "\t")

#Select Discount tire File
disc_crawl_path=filedialog.askopenfilename(title='Choose Discount Tire Output ')
disc_crawl=pd.read_excel(disc_crawl_path)

#Sams Club File
sams_crawl_path=filedialog.askopenfilename(title='Choose Samsclub Output')
sams_crawl=pd.read_csv(sams_crawl_path, sep = "\t")

#Status File(Tires Costco)
status_crawl_path=filedialog.askopenfilename(title='Choose the Tirescostco file for Status')
status_crawl=pd.read_csv(status_crawl_path, sep = "\t")

#BJ's Function
bj_file.columns = bj_file.columns.str.replace("'","")
bj_file.rename({'Unnamed: 7': 'BJ_price'}, axis =1 , inplace= True)
bj_file['BJ_price']=bj_file['BJs Price']
bj_file.columns=bj_file.columns.to_flat_index()
    
bj_crawl['x']=bj_crawl['Buybox_Winner_Vendor_Name'].astype(str)+' '+bj_crawl['Buybox_Winner_Vendor_Price'].astype(str)+' '+bj_crawl['Top_3_Buybox_Winners'].astype(str)
bj_file.rename(columns = {'Item No.':'ebags_unique_identifier'}, inplace=True)
    
df_1= pd.merge(bj_file, bj_crawl [['ebags_unique_identifier','product_id','product_name','x','Regular_price','product_url']], on='ebags_unique_identifier', how='left')
df_1.rename(columns = {'ebags_unique_identifier':'Item_No.'}, inplace=True)
df_1.rename({'Regular_price': 'X'}, axis =1 , inplace= True)

bj_out=df_1
    
bj_out["X"].fillna("na", inplace = True)
bj_out.loc[bj_out.X == "na", "Y"] = "na"
bj_out.loc[bj_out.Y == "na", "BJ_price"] = "Unique to Costco"
  
bj_out['Y']=np.where((bj_out['BJ_price']==bj_out['X']) & (bj_out['City']=="Brookfield"),True, False)
    
bj_out['BJ_price']=np.where(((bj_out['Y']==False) & (bj_out['X']!="na")),bj_out['X'],bj_out['BJ_price'])
    
bj_out['Product No.']=np.where((bj_out['Product No.'].notnull()),bj_out['product_id'],bj_out['Product No.'])
   
bj_out['BJs Price']=np.where(((bj_out["City"]=="Brookfield") & (bj_out['BJ_price']!="Unique to Costco")),bj_out['BJ_price'],bj_out['BJs Price']) 
bj_out['BJs Price']=np.where(((bj_out["City"]=="Brookfield") & (bj_out['BJ_price']=="Unique to Costco")),bj_out['BJ_price'],bj_out['BJs Price']) 

del bj_out['BJ_price']
    
bj_out['Size + Speed Rating + Load Rating.1']=bj_out['x']
bj_out['Product Name.1']=bj_out['product_name']
    
bj_out['Product No.']=bj_out['Product No.'].astype('Int64')
    
bj_out['product_url']=np.where(((bj_out['Y']==False) & (bj_out.X.isnull())),bj_out['product_url'],'NA')

bj_out['link']=np.where((bj_out['product_url']=='NA'),'https://www.example.com',bj_out['product_url'])

words = 'Sorry, no tire can be found'
for i in bj_out["link"]:
    if "tires.bjs.com" in i:
        for i in bj_out['link']:
            temp = urllib.request.urlopen(i)
            HTML = temp.read().decode("utf-8")
            if words in HTML:
                bj_out['BJs Price']= 'Unique to Costco'
                    
bj_out.columns

try:
    del bj_out["link"]
except:
    pass

try:              
    del bj_out["product_url"]
except:
    pass
              
try:
    del bj_out["X"]    
except:
    pass
          
try:
    del bj_out["x"] 
except:
    pass
             
try:
    del bj_out["product_name"]     
except:
    pass
         
try:
    del bj_out["Unnamed: 12"]    
except:
    pass
          
try:
    del bj_out["Y"]    
except:
    pass
   
try:
    del bj_out["product_id"]    
except:
    pass     
                
bj_final_out = bj_out.copy()
bj_final_out["BJs Price"]=np.where((bj_final_out["City"]=="ECOM"),"NA",bj_final_out["BJs Price"])

bj_final_out["Product No."]=np.where((bj_final_out["City"]=="ECOM"),"",bj_final_out["Product No."])
bj_final_out["Product Name.1"]=np.where((bj_final_out["City"]=="ECOM"),"",bj_final_out["Product Name.1"])
bj_final_out["Size + Speed Rating + Load Rating.1"]=np.where((bj_final_out["City"]=="ECOM"),"",bj_final_out["Size + Speed Rating + Load Rating.1"])
   
bj_final_out["Product No."]=np.where((bj_final_out["BJs Price"]=="Unique to Costco"),"Unique to Costco",bj_final_out["Product No."])
bj_final_out["Product Name.1"]=np.where((bj_final_out["BJs Price"]=="Unique to Costco"),"Unique to Costco",bj_final_out["Product Name.1"])
bj_final_out["Size + Speed Rating + Load Rating.1"]=np.where((bj_final_out["BJs Price"]=="Unique to Costco"),"Unique to Costco",bj_final_out["Size + Speed Rating + Load Rating.1"])
   

                     

#Discount Tire Function
disc_crawl["price"].fillna("0", inplace= True)
disc_file.rename({'Unnamed: 7': 'Price'}, axis =1 , inplace= True)
disc_file['Price']=disc_file['Discount Tire Price']

disc_file.columns=disc_file.columns.to_flat_index()

disc_file['z']=disc_file['Item No.'].astype(str)+''+disc_file['ZIP'].astype(str)
disc_crawl['z']=disc_crawl['ebags'].astype(str)+''+disc_crawl['zip_code'].astype(str)
df_2= pd.merge(disc_file, disc_crawl [['z','price','size','product_name','product_id']], on='z', how='left')
df_2.rename({'price': 'X'}, axis =1 , inplace= True)
disc_out=df_2
disc_out["X"].fillna("na", inplace = True)
#disc_out.loc[disc_out.X == "na", "Y"] = "na"



disc_out['Y']=np.where((disc_out['Price']==disc_out['X']) ,True, False)
#disc_out['Y']=np.where((disc_out["X"]=='na'),'na',disc_out["Y"])

#disc_out["Price"].fillna("na", inplace = True)

disc_out["size"].fillna("na", inplace = True)
disc_out["product_name"].fillna("na", inplace = True)
disc_out["product_id"].fillna("na", inplace = True)

disc_out["Size + Speed Rating + Load Rating.1"]=np.where(((disc_out['Y']==False)&(disc_out['Price']=="Unique to Costco")&(disc_out['X']!="0")&(disc_out['X']!="na")),disc_out["size"],disc_out["Size + Speed Rating + Load Rating.1"])
disc_out["Product No."]=np.where(((disc_out['Y']==False)&(disc_out['Price']=="Unique to Costco")&(disc_out['X']!="0")&(disc_out['X']!="na")),disc_out["product_id"],disc_out["Product No."])
disc_out['Product Name.1']=np.where(((disc_out['Y']==False)&(disc_out['Price']=="Unique to Costco")&(disc_out['X']!="0")&(disc_out['X']!="na")),disc_out['product_name'], disc_out['Product Name.1'] )

disc_out['Product Name.1']=np.where(((disc_out['Y']==False)&(disc_out["X"]=='0')),"Unique to Costco", disc_out['Product Name.1'] )
disc_out['Product No.']=np.where(((disc_out['Y']==False)&(disc_out["X"]=='0')&(disc_out['Price']=="Unique to Costco")),"Unique to Costco", disc_out['Product No.'] )
disc_out['Size + Speed Rating + Load Rating.1']=np.where(((disc_out['Y']==False)&(disc_out["X"]=='0')&(disc_out['Price']=="Unique to Costco")),"Unique to Costco", disc_out['Size + Speed Rating + Load Rating.1'] )

disc_out['Price']=np.where(((disc_out['Y']==False)&(disc_out['X']!="na")&(disc_out['X']!="0")),disc_out['X'], disc_out['Price'])
#disc_out['Price']=np.where(((disc_out['Y']==False)&(disc_out['X']!='Unique to Costco')&(disc_out["X"]=="na")),disc_out['Price'], disc_out['X'])
disc_out['Price']=np.where(((disc_out['Y']==False)&(disc_out['X']=="0")),"Unique to Costco",disc_out['Price'])


disc_out["Discount Tire Price"]=disc_out["Price"].copy()

del disc_out["Price"]
disc_out.columns

try:
    del disc_out["Y"]
except:
    pass

try:
    del disc_out["product_id"]              
except:
    pass

try:
    del disc_out["product_url"]              
except:
    pass

try:
    del disc_out["product_name"]
except:
    pass

try:
    del disc_out["size"]
except:
    pass

try:              
    del disc_out["X"]
except:
    pass
              
try:
    del disc_out["z"]    
except:
    pass
          
try:
    del disc_out["Unnamed: 16"] 
except:
    pass
             
try:
    del disc_out["Unnamed: 15"]     
except:
    pass
         
try:
    del disc_out["Unnamed: 14"]    
except:
    pass

try:
    del disc_out["Unnamed: 13"]    
except:
    pass

try:
    del disc_out["Unnamed: 12"]    
except:
    pass

disc_final_out = disc_out.copy()

disc_final_out=disc_final_out.fillna("n/a")

disc_final_out["Discount Tire Price"]=np.where((disc_final_out["City"]=="ECOM"),"n/a",disc_final_out["Discount Tire Price"])

disc_final_out["Product No."]=np.where((disc_final_out["City"]=="ECOM"),"",disc_final_out["Product No."])
disc_final_out["Product Name.1"]=np.where((disc_final_out["City"]=="ECOM"),"",disc_final_out["Product Name.1"])
disc_final_out["Size + Speed Rating + Load Rating.1"]=np.where((disc_final_out["City"]=="ECOM"),"",disc_final_out["Size + Speed Rating + Load Rating.1"])
disc_final_out["1 = different Costco and Discount product name"]=np.where((disc_final_out["City"]=="ECOM"),"",disc_final_out["1 = different Costco and Discount product name"])


#Sams Club Function
sams_file.columns = sams_file.columns.str.replace("'","")

sams_crawl["InStore_Regular_Price"]=np.where((sams_crawl["price2"]=="Member Only Price"),sams_crawl["Online_Markdown_Price"],sams_crawl["InStore_Regular_Price"])

sams_file['SamsPrice']=sams_file['Sams Club Price']

sams_file.columns=sams_file.columns.to_flat_index()

sams_file.rename(columns = {'Item No.':'ebags_unique_identifier'}, inplace=True)
df_3= pd.merge(sams_file, sams_crawl [['ebags_unique_identifier','InStore_Regular_Price','SKU_No','product_name','Product_Url','Size']], on='ebags_unique_identifier', how='left')
df_3.rename(columns = {'ebags_unique_identifier':'Item_No.'}, inplace=True)
df_3.rename({'InStore_Regular_Price': 'X'}, axis =1 , inplace= True)
sams_out=df_3
sams_out["X"].fillna("na", inplace = True)

sams_out['Y']=np.where((sams_out['SamsPrice']==sams_out['X']) & (sams_out['City']=="ALBUQUERQUE"),True, False)

sams_out.loc[sams_out.X == "na", "Y"] = "na"
sams_out.loc[sams_out.Y == "na", "SamsPrice"] = "Unique to Costco"


#sams_out['SamsPrice']=np.where(((sams_out['Y']==False) & (sams_out['X']!="na")),sams_out['X'],sams_out['SamsPrice'])
sams_out['SamsPrice']=np.where((sams_out['Y']==False),sams_out['X'],sams_out['SamsPrice'])
sams_out['SamsPrice']=np.where((sams_out['Y']=='na'),"Unique to Costco",sams_out['SamsPrice'])

sams_out['Product No.']=np.where(((sams_out['Product No.'].notnull()) & (sams_out["City"]=="ALBUQUERQUE")),sams_out['SKU_No'],sams_out['Product No.'])


sams_out['Size + Speed Rating + Load Rating.1']=np.where((sams_out['Y']==False),sams_out['Size'],sams_out["Size + Speed Rating + Load Rating.1"])

#sams_out['Sams Club Price']=np.where(((sams_out["City"]=="ALBUQUERQUE") & (sams_out['SamsPrice']!="Unique to Costco")),sams_out['SamsPrice'],sams_out['Sams Club Price']) 
sams_out['Sams Club Price']=sams_out["SamsPrice"].copy()


sams_out['Product No.']=sams_out['SKU_No']
sams_out['Product Name.1']=sams_out['product_name']

sams_out['Product No.']=sams_out['Product No.'].astype('Int64')

sams_out['Product_Url']=np.where(((sams_out['Y']==False) & (sams_out.X.isnull())),sams_out['Product_Url'],'NA')

sams_out['link']=np.where((sams_out['Product_Url']=='NA'),'https://www.example.com',sams_out['Product_Url'])

words = 'Sorry, no tire can be found'
for i in sams_out["link"]:
    if "tires.bjs.com" in i:
        for i in sams_out['link']:
            temp = urllib.request.urlopen(i)
            HTML = temp.read().decode("utf-8")
            if words in HTML:
                sams_out['Sams Club Price']= 'Unique to Costco'

sams_out.columns

try:
    del sams_out["link"]
except:
    pass

try:
    del sams_out["Y"]
except:
    pass

try:              
    del sams_out["Size"]
except:
    pass
              
try:
    del sams_out["Product_Url"]    
except:
    pass
          
try:
    del sams_out["product_name"] 
except:
    pass
             
try:
    del sams_out["SKU_No"]     
except:
    pass
         
try:
    del sams_out["X"]    
except:
    pass

try:
    del sams_out["SamsPrice"]    
except:
    pass

try:
    del sams_out["Unnamed: 12"]    
except:
    pass

try:
    del sams_out["Unnamed: 7"]    
except:
    pass


sams_final_out = sams_out.copy()

#del sams_final_out['Unnamed: 7']
sams_final_out['Product No.']=np.where((sams_final_out['Sams Club Price']=="Unique to Costco"),"Unique to Costco",sams_final_out['Product No.'])
sams_final_out['Product Name.1']=np.where((sams_final_out['Sams Club Price']=="Unique to Costco"),"Unique to Costco",sams_final_out['Product Name.1'])
sams_final_out['Size + Speed Rating + Load Rating.1']=np.where((sams_final_out['Sams Club Price']=="Unique to Costco"),"Unique to Costco",sams_final_out['Size + Speed Rating + Load Rating.1'])

sams_final_out=sams_final_out.fillna("n/a")

sams_final_out["Sams Club Price"]=np.where((sams_final_out["City"]=="ECOM"),"NA",sams_final_out["Sams Club Price"])

sams_final_out["Product No."]=np.where((sams_final_out["City"]=="ECOM"),"",sams_final_out["Product No."])
sams_final_out["Product Name.1"]=np.where((sams_final_out["City"]=="ECOM"),"",sams_final_out["Product Name.1"])
sams_final_out["Size + Speed Rating + Load Rating.1"]=np.where((sams_final_out["City"]=="ECOM"),"",sams_final_out["Size + Speed Rating + Load Rating.1"])
sams_final_out["1 = different Costco and BJs product name"]=np.where((sams_final_out["City"]=="ECOM"),"",sams_final_out["1 = different Costco and BJs product name"])

#sams_final_out=sams_final_out.iloc[: ,:-1]

#Cover page
cvr.rename(columns = {'Item Number':'Item No.'}, inplace=True)

newdisc=disc_final_out.loc[disc_final_out['City'] != "ECOM"].copy()
#newdisc["Discount Tire Price"]=newdisc["Discount Tire Price"].astype(str)
    
result_df = newdisc.groupby('Item No.').agg({'Discount Tire Price': ['min', 'max']})
result_df["Min_price"]=result_df["Discount Tire Price"]["min"]
result_df["Max_price"]=result_df["Discount Tire Price"]["max"]
result_df.reset_index(inplace=True)
result_df.columns = result_df.columns.droplevel(1)

#result_df["Min_price"]=np.where(((result_df["Min_price"])<=(result_df["Max_price"])),result_df["Min_price"],result_df["Max_price"])
#result_df["Max_price"]=np.where(((result_df["Min_price"])<=(result_df["Max_price"])),result_df["Max_price"],result_df["Min_price"])


del result_df["Discount Tire Price"]

df_4= pd.merge(cvr, result_df [['Item No.','Min_price','Max_price']], on='Item No.', how='left')

cvr["High.3"]=df_4["Max_price"].copy()
cvr["Low.3"]=df_4["Min_price"].copy()

#cvr["hi"]=cvr["High.3"].copy()
#cvr["lo"]= cvr["Low.3"].copy()


cvr.rename(columns = {'Item No.':'Item_No.'}, inplace=True)

newsam=sams_final_out.loc[sams_final_out['City'] != "ECOM"]

df_5= pd.merge(cvr, newsam [['Item_No.','Sams Club Price']], on='Item_No.', how='left')

cvr["High.2"] = df_5["Sams Club Price"]
cvr["Low.2"] = df_5["Sams Club Price"]

newbj=bj_final_out.loc[bj_final_out['City'] != "ECOM"]

df_6= pd.merge(cvr, newbj [['Item_No.','BJs Price']], on='Item_No.', how='left')

cvr["High.1"] = df_6["BJs Price"]
cvr["Low.1"] = df_6["BJs Price"]

cvr.rename(columns = {'Item_No.':'Item Number'}, inplace=True)

cvr=cvr.fillna("n/a")


#Status
status.insert(loc = 3, column = 'date', value = "")
status_crawl['new']=status_crawl['product_id']
status.rename(columns = {'ITEM NO':'product_id'}, inplace=True)

df_stat= pd.merge(status, status_crawl [['product_id','new']], on='product_id', how='left')
df_stat.rename(columns = {'product_id':'ITEM NO'}, inplace=True)
status=df_stat.copy()

status["new"].fillna(0, inplace = True)
status['new']=status['new'].astype('int64')
status["date"]=status["new"]

status['date']=np.where((status['date']==0) ,"na",status['date'])

status['date']=np.where((status['date']=="na") ,status['date'] ,"Online")

status.rename({'date': datetime.datetime.now().date()}, axis =1 , inplace= True)

del status['new']

status=status.fillna("n/a")

status.iloc[ : , 3 ]=np.where((status.iloc[ : , 3 ]=="na"),"Offline",status.iloc[ : , 3 ])
#########################################################################################################
#ECOM Changes
cvr.rename(columns = {'Item Number':'Item_No.'}, inplace=True)
#BJ's
bj_final_out= pd.merge(bj_final_out, cvr [['Item_No.','Costco Sell Price']], on='Item_No.', how='left')
bj_final_out["BJs Price"]= np.where((bj_final_out["City"]=="ECOM"),bj_final_out["Costco Sell Price"],bj_final_out["BJs Price"])
bj_final_out.columns
del bj_final_out["Costco Sell Price"]
#Discount Tire
cvr.rename(columns = {'Item_No.':'Item No.'}, inplace=True)

disc_final_out= pd.merge(disc_final_out, cvr [['Item No.','Costco Sell Price']], on='Item No.', how='left')
disc_final_out["Discount Tire Price"]= np.where((disc_final_out["City"]=="ECOM"),disc_final_out["Costco Sell Price"],disc_final_out["Discount Tire Price"])
disc_final_out.columns
del disc_final_out["Costco Sell Price"]
#Sams Club
cvr.rename(columns = {'Item No.':'Item_No.'}, inplace=True)

sams_final_out= pd.merge(sams_final_out, cvr [['Item_No.','Costco Sell Price']], on='Item_No.', how='left')
sams_final_out["Sams Club Price"]= np.where((sams_final_out["City"]=="ECOM"),sams_final_out["Costco Sell Price"],sams_final_out["Sams Club Price"])
sams_final_out.columns
del sams_final_out["Costco Sell Price"]


cvr.rename(columns = {'Item_No.':'Item Number'}, inplace=True)


#Price Difference

cvr.insert(loc = 5, column = 'Price_Difference', value = "")


cvr["Low.1"]=cvr["Low.1"].replace({'Unique to Costco': 9999999})
cvr["Low.2"]=cvr["Low.2"].replace({'Unique to Costco': 9999999})
cvr["Low.3"]=cvr["Low.3"].replace({'Unique to Costco': 9999999})
cvr["High.3"]=cvr["High.3"].replace({'Unique to Costco': 9999999})
#cvr["Costco Sell Price"]=cvr["Costco Sell Price"].replace({'unique to Costco': 9999999})
cvr["Costco Sell Price"]=cvr["Costco Sell Price"].replace({'n/a': 0})


cvr["Low.1"]=cvr["Low.1"].astype(float)
cvr["Low.2"]=cvr["Low.2"].astype(float)
cvr["High.3"]=cvr["High.3"].astype(float)
cvr["Low.3"]=cvr["Low.3"].astype(float)
cvr["Costco Sell Price"]=cvr["Costco Sell Price"].astype(float)

cvr["min"] = cvr[['Low.1','Low.2','High.3','Low.3']].min(axis=1)

cvr["Price_Difference"]=cvr["Costco Sell Price"]-cvr["min"]

cvr["Price_Difference"]=cvr["Price_Difference"].astype(float)
cvr["Price_Difference"]=(round(cvr["Price_Difference"], 2))

cvr["Price_Difference"]=np.where(((cvr["Price_Difference"]>99999)|(cvr["Price_Difference"]<-999)),"Unique to Costco",cvr["Price_Difference"])
cvr["Price_Difference"]=np.where((cvr["Costco Sell Price"]==0),'n/a',cvr["Price_Difference"])

cvr["Costco Sell Price"]=cvr["Costco Sell Price"].replace({0: 'n/a'})

cvr["Low.1"]=cvr["Low.1"].replace({9999999: 'Unique to Costco' })
cvr["Low.2"]=cvr["Low.2"].replace({9999999: 'Unique to Costco'})
cvr["Low.3"]=cvr["Low.3"].replace({9999999: 'Unique to Costco'})
cvr["High.3"]=cvr["High.3"].replace({9999999: 'Unique to Costco'})

cvr.columns
#cvr=cvr.drop(cvr.columns[6],axis=1)
del cvr["Price\nDifference"]
del cvr["min"]

output_path = filedialog.asksaveasfilename(title="Choose the name and directory to save file", defaultextension=".xlsx",filetypes=[("Excel files", "*.xlsx"),("All files", "*.*") ])
writer = pd.ExcelWriter(output_path, engine = 'xlsxwriter')
cvr.to_excel(writer, sheet_name = "Cover Page", index=False)
summary.to_excel(writer, sheet_name = "Summary", index=False)
status.to_excel(writer, sheet_name = "Status-By-Date", index=False)
bj_final_out.to_excel(writer, sheet_name = "TOP 200 - BJ's", index=False)
disc_final_out.to_excel(writer, sheet_name = "TOP 200 - Discount Tire", index=False)
sams_final_out.to_excel(writer, sheet_name = "TOP 200 - Sam's Club", index=False)
writer.save()
writer.close()

root.destroy()




