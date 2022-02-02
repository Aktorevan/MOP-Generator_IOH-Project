from mailmerge import MailMerge
import pandas as pd
from datetime import date
import os

#To load the database
data = pd.read_excel('db.xlsx')

#To return the list of MOP file name
filterMOP = data["MOP File Name"].unique()

#Template .docx
base_temp = os.path.basename("Template")

reroute_temp = base_temp+'/dis_route.docx' #Dismantle and reoute 
software_temp = base_temp+'/software.docx' #Software upgrade  
frequency_temp = base_temp+'/frequency.docx' #Change frequency
cutover_temp = base_temp+'/cutover.docx' #Cutover activity
#vlan_temp = '' #Vlan ID upgrade
#ts_temp = '' #Troublshooting activity

#Output folder
base_out = os.path.basename("Output")

def proc_MOP(template):

    '''
    The function is used to process the MOP with the spesific template, based on the SOW 

    '''

    #To load the template
    document = MailMerge(template)

    today = date.today() #Today's date written on the first page
    title = str(i)       #MOP Title
    num_impact = len(filterTable.index) #The number of DUID list
   
    #Parsing process with dictionary for creating table
    tableRows = []
    j=0

    while j < num_impact:
      
      region_name = filterTable['NE Region'].iloc[j] #To exctract the spesific region
      exec_date = filterTable['Execution Date'].iloc[j] #To extract the spesific date
      exec_time = filterTable['Time'].iloc[j] #To extract the spesific time
      
      filterDUID = filterTable['Relative NE '].iloc[j] #Parsing each DUID
      filterQty = filterTable['Dependency Qty'].iloc[j] #Parsing each qty
      filterDependency = filterTable['Site Dependency'].iloc[j] #Parsing the dependency list
      filterDataSource = filterTable['Impact Data Source'].iloc[j] #Parsing data source

      #Merging list of 'tableRows' with the dictionary from each parsed item
      tableRows.append({'duid' : str(filterDUID), 
                      'qty' : str(filterQty), 
                      'list' : str(filterDependency), 
                      'source' : str(filterDataSource)})

      j +=  1


    document.merge_rows('duid',tableRows)

    document.merge(predate = str(today.strftime("%B %d, %Y")), 
                   titlemop = title,
                   linknum = str(num_impact),
                   region = str(region_name),
                   date = str(exec_date.strftime("%d %B %Y")),
                   time = str(exec_time),        
                   duid = tableRows)

    #Export to output folder
    document.write((base_out+'/{0}.docx' .format(title)))


def proc_DEP(name):

    db = data[["Relative NE ", "Site Dependency", 'MOP File Name']]

    filMOP = db[db['MOP File Name'] == name][["Relative NE ", "Site Dependency", 'MOP File Name']]
    filMOP.rename({"Site Dependency" : "Site Id"}, axis = 1, inplace = True)
    filMOP["Site Id"] = filMOP["Site Id"].str.split(",") #Seperate values from commas
    filMOP = filMOP["Site Id"].explode()
    filMOP = filMOP.reset_index(drop = True) #To reset index
    filMOP = pd.DataFrame(filMOP) #Convert from Series to Dataframe
    filMOP = filMOP[filMOP["Site Id"] != "-"] 

    filMOP.insert(1, "NE Name", '') #Insert blank name entitled "NE Name" after column "Site Id"
      
    filMOP.to_excel(base_out+'/{}.xlsx'.format(name), sheet_name = "Site Impact", index = False)
    

for i in filterMOP:

  #Select the column
  filterTable = data[data["MOP File Name"] == str(i)][['Relative NE ','Scope', 'NE Region','Time', 'Execution Date', 'Dependency Qty', 'Site Dependency', 'Impact Data Source']]
  
  #Filter based on SOW for selecting the template
  filterSOW = filterTable['Scope'].unique()

  #Generating based on SOW
  if filterSOW == "MW Reroute":

    proc_MOP(reroute_temp)
    proc_DEP(i)
    print(str(i)+" generated")

  elif filterSOW == "Software Upgrade":

    proc_MOP(software_temp)
    proc_DEP(i)
    print(str(i)+" generated")
  
  elif filterSOW == "Change Frequency":

    proc_MOP(frequency_temp)
    proc_DEP(i)
    print(str(i)+" generated")

  elif filterSOW == "Cutover":

    proc_MOP(cutover_temp)
    proc_DEP(i)
    print(str(i)+" generated")
  

