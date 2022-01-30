from mailmerge import MailMerge
import pandas as pd
from datetime import date

#To load the database
data = pd.read_excel('C:/Users/awx1014609/Desktop/My Lab/sample_db.xlsx')

#To return the list of MOP file name
filterMOP = data["MOP File Name"].unique()


for i in filterMOP:

  #To load the template
  document = MailMerge('C:/Users/awx1014609/Desktop/My Lab/Template/dis_route.docx')

  #Select the column
  filterTable= data[data["MOP File Name"] == str(i)][['Relative NE ','NE Region','Time', 'Execution Date', 'Dependency Qty', 'Site Dependency', 'Impact Data Source']]

  today = date.today() #Today's date written on the first page
  title = str(i)       #MOP Title
  num_impact = len(filterTable.index) #The number of DUID list
 

  tableRows = []

  j=0

  while j < num_impact:
    
    region_name = filterTable['NE Region'].iloc[j] #To exctract the spesific region
    exec_date = filterTable['Execution Date'].iloc[j]
    exec_time = filterTable['Time'].iloc[j]
    
    filterDUID = filterTable['Relative NE '].iloc[j] #To exctract the spesific DUID
    filterQty = filterTable['Dependency Qty'].iloc[j] #To exctract the spesific qty
    filterDependency = filterTable['Site Dependency'].iloc[j] #To exctract the spesific dependency list
    filterDataSource = filterTable['Impact Data Source'].iloc[j] #To exctract the spesific data source


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

    
  
  
  

