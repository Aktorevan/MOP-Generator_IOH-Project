from mailmerge import MailMerge
import pandas as pd
from datetime import date

#To load the database
data = pd.read_excel('C:/Users/awx1014609/Desktop/My Lab/sample_db.xlsx')

#To return the list of MOP file name
filterMOP = data["MOP File Name"].unique()

#Template .docx
reroute_temp = 'C:/Users/awx1014609/Desktop/My Lab/Template/dis_route.docx' #Dismantle and reoute 
software_temp = 'C:/Users/awx1014609/Desktop/My Lab/Template/software.docx' #Software upgrade  
frequency_temp = 'C:/Users/awx1014609/Desktop/My Lab/Template/frequency.docx' #Change frequency
cutover_temp = 'C:/Users/awx1014609/Desktop/My Lab/Template/cutover.docx' #Cutover activity
#vlan_temp = '' #Vlan ID upgrade
#ts_temp = '' #Troublshooting activity


def proc_MOP(template):

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

    document.write(('C:/Users/awx1014609/Desktop/My Lab/Output/{0}.docx'.format(title)))

    

for i in filterMOP:

  #Select the column
  filterTable = data[data["MOP File Name"] == str(i)][['Relative NE ','Scope', 'NE Region','Time', 'Execution Date', 'Dependency Qty', 'Site Dependency', 'Impact Data Source']]
  
  #Filter based on SOW for selecting the template
  filterSOW = filterTable['Scope'].unique()

  #Generating based on SOW
  if filterSOW == "MW Reroute":

    proc_MOP(reroute_temp)
    print(str(i)+".docx generated")

  elif filterSOW == "Software Upgrade":

    proc_MOP(software_temp)
    print(str(i)+".docx generated")
  
  elif filterSOW == "Change Frequency":

    proc_MOP(frequency_temp)
    print(str(i)+".docx generated")

  elif filterSOW == "Cutover":

    proc_MOP(cutover_temp)
    print(str(i)+".docx generated")
  

