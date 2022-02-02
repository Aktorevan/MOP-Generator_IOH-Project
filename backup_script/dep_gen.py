from openpyxl import load_workbook
import pandas as pd
import os

#To load database
df = pd.read_excel('db.xlsx')
db = df[["Relative NE ", "Site Dependency", 'MOP File Name']]

#Output folder
base_out = os.path.basename("Output")

mop_list = df['MOP File Name'].unique()

for name in mop_list:
    
    filMOP = db[db['MOP File Name'] == name][["Relative NE ", "Site Dependency", 'MOP File Name']]
    filMOP.rename({"Site Dependency" : "Site Id"}, axis = 1, inplace = True)
    filMOP["Site Id"] = filMOP["Site Id"].str.split(",") #Seperate values from commas
    filMOP = filMOP["Site Id"].explode()
    filMOP = filMOP.reset_index(drop = True) #To reset index
    filMOP = pd.DataFrame(filMOP) #Convert from Series to Dataframe
    filMOP = filMOP[filMOP["Site Id"] != "-"] 

    filMOP.insert(1, "NE Name", '') #Insert blank name entitled "NE Name" after column "Site Id"
    
    filMOP.to_excel(base_out+'/{}.xlsx'.format(name), sheet_name = "Site Impact", index = False)