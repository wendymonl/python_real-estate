# -*- coding: utf-8 -*-
"""
Created on Tue Jul 23 12:29:57 2019

@author: monta
"""

from xlwt import Workbook
opentxtfile=open("LaredoParks.txt") #all real estate parks in Laredo are in this file
parkslines=opentxtfile.readlines()

workbookname="Laredo Industrial Parks" #this workbook will house each industrial park in Laredo. 
workbookname= Workbook()
numberofparks=len(parkslines)

#each industrial park will be a sheet, and every building in the park will written into that sheet
for i in range(0,numberofparks-1):
    parknamewithenter=parkslines[i]
    parkname=parknamewithenter.split()[0]
    sheet1 = workbookname.add_sheet(parkslines[i])
    print(parkname+' GeographicIDs.txt')
    nameoffile=parkname+' GeographicIDs.txt'
    print(nameoffile)
    opengeographicidtxt=open(nameoffile) 
    geographicid_list=opengeographicidtxt.readlines() #include Geographic IDs of buildings for sorting later
