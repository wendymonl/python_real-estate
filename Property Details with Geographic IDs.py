# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 10:12:12 2019

@author: monta
"""

parkname="Laredo Dist."
import xlwt
from xlwt import Workbook
# Workbook is created 
workbookname=parkname+' Workbook'
workbookname= Workbook()
# add_sheet is used to create sheet. 
sheet1 = workbookname.add_sheet("Webb CAD data")
sheet1.write(1,0, "Input Geographic ID")
sheet1.write(1,3, "Building Street Address")
sheet1.write(1,1, "Property ID")
sheet1.write(1,2, "Scraped Geographic ID")
sheet1.write(1,4, "Reported Building Owner")
sheet1.write(1,5, "Building Owner Company")
sheet1.write(1,6, "Owner Address")
sheet1.write(1,7, "Owner City")
sheet1.write(1,8, "Owner State")
sheet1.write(1,9, "Zip Code")
sheet1.write(1,10,"Contact Name")
filename=parkname+" GeographicIDs.txt"
openparkfile=open(filename)
parkfilelines=openparkfile.readlines()
parkfile_length=len(parkfilelines)
for i in range (0,parkfile_length):
    geographicidindex=str(parkfilelines[i])
    geographicidindex_split=geographicidindex.split()[0]
    #will start putting data on on row 2 
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
    if "amp; " in ownername:
        import html
        ownername=html.unescape(ownername)
    style=xlwt.XFStyle()
    font=xlwt.Font()
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
        lenownername=len(ownername.split())
        font.bold=True
    style.font=font
    ownermailingaddresscode=lines[26] #this is the mailing address in html
    ownermailingaddress=ownermailingaddresscode.split("</td><td>")[1]
    ownermailingaddress_list=ownermailingaddress.split(" <br> ")
    lenownermailingadress_list=len(ownermailingaddress_list)
    ownercitystatezip=ownermailingaddress_list[lenownermailingadress_list-1]
    ownerstreetaddress=ownermailingaddress_list[lenownermailingadress_list-2]
    if lenownermailingadress_list is 3:
        ownercontactname=ownermailingaddress_list[lenownermailingadress_list-3]
    elif lenownername is not 0:
        ownercontactname=ownername
    else:
        ownercontactname=" "
    ownercitystateszip_list=ownercitystatezip.split()
    lenownercitystatezip=len(ownercitystatezip.split())
    ownerzip=ownercitystatezip.split()[lenownercitystatezip-1]
    ownerstate=ownercitystatezip.split()[lenownercitystatezip-2]
    ownercity=ownercitystatezip.split(",")[0]
    #put info into an excel sheet one row at a time
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