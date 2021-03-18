# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 10:12:12 2019

@author: monta
"""

#After creating a workbook of each Laredo Industrial Park, we will write the data into it
parkname="Laredo Dist."
import xlwt
from xlwt import Workbook
# Workbook is created 
workbookname=parkname+' Workbook'
workbookname= Workbook()
# add_sheet is used to create sheet. 

#write titles into the spreadsheet (this is the specific info we are interested in for each building.)
sheet1 = workbookname.add_sheet("Webb CAD data")
sheet1.write(1,0, "Input Geographic ID")
sheet1.write(1,3, "Building Street Address")
sheet1.write(1,1, "Property ID") #allows us to search on GIS system easily
sheet1.write(1,2, "Scraped Geographic ID") #we can later check for a match with "Input Geographic ID" to ensure data is accurate for each building
sheet1.write(1,4, "Reported Building Owner") #we are interested in the contact information of this owner
sheet1.write(1,5, "Building Owner Company") #We are interested in finding who the "big players" are in this real estate market
#store mailing address, city, state, zip and contact name in seperate columns.
#this method allows for easy printing of mailing labels later
sheet1.write(1,6, "Owner Address") 
sheet1.write(1,7, "Owner City")
sheet1.write(1,8, "Owner State")
sheet1.write(1,9, "Zip Code")
sheet1.write(1,10,"Contact Name")

#each industrial park has been stored as a .txt file containing a list of every Geogrphic ID of each industrial building in that park
filename=parkname+" GeographicIDs.txt"
openparkfile=open(filename)
parkfilelines=openparkfile.readlines()
parkfile_length=len(parkfilelines)

#populate info for each building by first converting html into readable text and numbers, and then write those properties in their respective rows and columns.
for i in range (0,parkfile_length):
    geographicidindex=str(parkfilelines[i])
    geographicidindex_split=geographicidindex.split()[0]
    #will start putting data on on row 2. Then fill one row at a time.
    rowinterger=i+2
    sheet1.write(rowinterger,0, geographicidindex_split)
    newname=geographicidindex_split+".txt"
    opentxtfile=open(newname)
    lines=opentxtfile.readlines()
    if len(lines) is not 31:
        print("error in ", newname, " data")
        
    #scrape property id intergers
    propertyidcode=lines[4] #this is the code in html
    propertyid_split=propertyidcode.split("</td>")
    propertyid=propertyid_split[1].split("<td>")[1]
    
    #scrape geographic id
    geographicidcode=lines[6] #this is the code in html
    geographicid=geographicidcode.split("</td><td>")[1]
    
    #scrape building address as string
    addresscode=lines[16] #this is the code in html
    fulladdress=addresscode.split("</td><td>")[1]
    streetadress=fulladdress.split("<br>")[0]
    
    #scrape owner information
    ownernamecode=lines[24] #this is the code for owner name in html
    ownername=ownernamecode.split("</td><td>")[1]
   
    #clean up the readability of owner names. Some buildings are owned by companies and some are owned by individuals.
    #we are interested in individual owners of buildings or key individuals in companies who own these buildings
    if "amp; " in ownername:
        import html
        ownername=html.unescape(ownername)
    style=xlwt.XFStyle()
    font=xlwt.Font()
    
    #types of companies who own buildings:
    if "LLC" in ownername:
        company=True
    if "LC" in ownername:
        company=True
    elif "LTD" in ownername:
        company=True
    elif " LP" or " L P" in ownername:
        company=True
    elif "PARTNERSHIP" in ownername:
        company=True
    elif "CORPORATION" in ownername:
        company=True
    elif " CO" in ownername:
        company=True
    elif "INC" in ownername:
        company=True
    elif "SERVICES" in ownername:
        company=True
    elif "ESTATE" in ownername:
        company=True
    elif "CITY OF" in ownername:
        company=True
    else:
        company=False
    if company is True:
        lenownername=0
        font.bold=False
    else:
        lenownername=len(ownername.split()) #split owner's name into first name and last name
        font.bold=True
    style.font=font
    
     #scrape the owner's mailing address. convert html code into string
    ownermailingaddresscode=lines[26] #this is the owner's mailing address in html
    ownermailingaddress=ownermailingaddresscode.split("</td><td>")[1]
    ownermailingaddress_list=ownermailingaddress.split(" <br> ")
    lenownermailingadress_list=len(ownermailingaddress_list) #mailing address as a list containing street address, city, state, and zip
    ownercitystatezip=ownermailingaddress_list[lenownermailingadress_list-1] #city state zip from mailing address
    ownerstreetaddress=ownermailingaddress_list[lenownermailingadress_list-2] #mailing street address
    
    if lenownermailingadress_list is 3: #this extracts the name of a contact person listed when a building is owned by a company
        ownercontactname=ownermailingaddress_list[lenownermailingadress_list-3]
    elif lenownername is not 0: #contact person is the owner when a building is owned by an individual
        ownercontactname=ownername
    else:
        ownercontactname=" " #Sometimes there is no contact person listed. 
        
    #now seperate the 'city, state, zip' array into individual objects. 
    ownercitystateszip_list=ownercitystatezip.split()
    lenownercitystatezip=len(ownercitystatezip.split()) 
    ownerzip=ownercitystatezip.split()[lenownercitystatezip-1] #mailing address zip code
    ownerstate=ownercitystatezip.split()[lenownercitystatezip-2] #mailing addresss state
    ownercity=ownercitystatezip.split(",")[0] #mailing addresss city
    
    #put info into an excel sheet one row at a time.
    sheet1.write(rowinterger,2, geographicid)
    sheet1.write(rowinterger,3, streetadress.title())
    sheet1.write(rowinterger,1, propertyid)
    sheet1.write(rowinterger,4, ownername,style=style)
    sheet1.write(rowinterger,6, ownerstreetaddress.title())
    sheet1.write(rowinterger,7, ownercity.title())
    sheet1.write(rowinterger,8, ownerstate)
    sheet1.write(rowinterger,9, ownerzip.split("-")[0])
    sheet1.write(rowinterger,10, ownercontactname.title())
    i=i+1
workbookname.save(parkname+'Webb CAD data.xls')  
