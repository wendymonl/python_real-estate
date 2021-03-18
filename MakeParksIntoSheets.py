# -*- coding: utf-8 -*-
"""
Created on Tue Jul 23 12:29:57 2019

@author: monta
"""

from xlwt import Workbook
opentxtfile=open("LaredoParks.txt")
parkslines=opentxtfile.readlines()
workbookname="Laredo Industrial Parks"
workbookname= Workbook()
numberofparks=len(parkslines)
for i in range(0,numberofparks-1):
    parknamewithenter=parkslines[i]
    parkname=parknamewithenter.split()[0]
    sheet1 = workbookname.add_sheet(parkslines[i])
    print(parkname+' GeographicIDs.txt')
    nameoffile=parkname+' GeographicIDs.txt'
    print(nameoffile)
    opengeographicidtxt=open(nameoffile)
    geographicid_list=opengeographicidtxt.readlines()