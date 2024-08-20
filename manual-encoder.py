
import pandas as pd
import xlrd, xlwt
from xlwt import Workbook

class Address:

    number = 0
    street = ""
    suffix = ""
    extra = ""
     

    def __init__(self, ad):
        ad = ad.split()

        try:
            ad[0] = int(ad[0])
        except ValueError:
            ad = input("Type standard address for " + ' '.join(ad) + ": ")


        self.ad = ad


def standardize(v):

    keys = {
        "STREET" : "ST",
        "ROAD" : "RD",
        "PARKWAY" : "PKWY",
        "DRIVE" : "DR",
        "AVENUE" : "AVE",
        "BOULEVARD" : "BLVD",
        "SUITE" : "STE",
        "APARTMENT" : "APT",
        "FLOOR" : "FL",
        "UNIT" : "UNIT",
        "ROOM" : "RM"
    }

    
    length = len(v.ad)

    if isinstance(v.ad, str):
        return v.ad

    for z in range(length):
        if keys.get(str(v.ad[z]).upper()) is not None:
            v.ad[z] = keys.get(v.ad[z].upper())

    rs = ""

    for q in range(length):
        rs = rs + str(v.ad[q]).upper() + " "
    
    return rs


addressCol = 5

datawb = xlrd.open_workbook('taa_2015_2017.xls')
datawb = datawb.sheet_by_index(0)

wb = Workbook()
fmt = wb.add_sheet('Formatted')
rfa = ["ADDRESS"]

for i in range(1, datawb.nrows):
    a = Address(datawb.cell_value(i,addressCol))
    rfa.append(standardize(a))


l = len(rfa)
for z in range(len(rfa)):
    fmt.write(z, 0, datawb.cell_value(z,0))
    fmt.write(z, 1, datawb.cell_value(z,1))
    fmt.write(z, 2, datawb.cell_value(z,2))
    fmt.write(z, 3, datawb.cell_value(z,3).upper())
    fmt.write(z, 4, datawb.cell_value(z,4).upper())
    fmt.write(z, 5, rfa[z])
    fmt.write(z, 6, datawb.cell_value(z,6).upper())
    fmt.write(z, 7, datawb.cell_value(z,7).upper())
    fmt.write(z, 8, datawb.cell_value(z,8))



wb.save('Formatted Addresses.xls')
