# Address Standardizer

difficult_addresses = []

def diff(fullAddress, Address): # This function standardizes addresses that the standardize() function finds difficult.
    r = Address[0].split()

    n = 0
    hwy = False
    count = 0
    for el in r:
        if el.isnumeric():
            n = n + 1
        if el.upper() == "HWY" or el.upper() == "HIGHWAY":
            hwy = True
        count = count + 1


    if r[0] + " " + r[1] == "PO Box" or r[0] + " " + r[1] == "P.O. Box": # This block standardizes addresses that only consist of a PO Box.
        return "PO BOX " + r[2]
    elif hwy and r[0].isnumeric():
        rt = ""
        for elm in r:
            if elm.upper() != "HWY" and elm.upper() != "HIGHWAY":
                rt = rt + elm.upper() + " "
            else:
                rt = rt + "HWY "
        return rt
    elif n >= 2: # Some addresses take the form of "## Road Name ##", where the road name contains a number. This block formats those.
        j = 0
        for el in r:
            if el.isnumeric():
                j = j + 1
            if j >= 2:
                save = el
                r.remove(el)
                Address[0] = " ".join(r)
                try:
                    return standardize(Address) + " " + save.upper()
                except TypeError:
                    pass
    
    difficult_addresses.append(fullAddress)
    return "UNFORMATTED" # All addresses that are too difficult for this function are returned as "UNFORMATTED"

def standardize(inp): # This function standardizes the majority (~90.6%) of the addresses.

    final_address = []
    Address = inp
    # PART #1---Split the Address into Relevant Parts

    import pyap
    full = (Address[0] + " " + Address[1] + " " + Address[2] + " " + Address[3]).upper()

    parsed = pyap.parse(full, country='US') # Using the pyap library, we can separate an address into its constituent parts.
    if len(parsed) == 0:
        return diff(full, Address)
    for ad in parsed:
        break
        print(ad)
        print(ad.as_dict())

    # PART #2---Standardize the Individual Parts

    sn = ad.as_dict()["street_number"] # This part standardizes the street number.
    from word2number import w2n
    if not sn.isnumeric():
        try:
            final_address.append(str(w2n.word_to_num(sn))) # Most street numbers are already in a standard form, except for those that have been spelled out (i.e "One")
        except ValueError: 
            return diff(full, Address)
    else:
        final_address.append(sn) 

    final_address.append(ad.as_dict()["street_name"]) # Street names do not need to be edited.

    st = ad.as_dict()["street_type"] # This block standardizes the street type (i.e. "Road," "Street," "Parkway.")
    work = xlrd.open_workbook("conventions.xls") # Accessing the list of USPS standard conventions, we can deduce what the street type ought to be converted to
    conventions = work.sheet_by_index(0)
    for q in range(5, conventions.nrows):
        z = q
        if conventions.cell_value(q, 1) == st:
            fin = conventions.cell_value(q, 2)
            while fin == "":
                z = z - 1    
                fin = conventions.cell_value(z, 2)
            final_address.append(fin)
            break

    if ad.as_dict()["occupancy"] != None: # This block standardizes the occupancy (i.e. "Suite ##")
        sp = ad.as_dict()["occupancy"].split(" ")
        if len(sp) > 1:
            final_address.append("STE " + sp[1])

    return ' '.join(final_address)

file_title = "taa_2015_2017.xls"
import xlrd, xlwt
from xlwt import *
wb = xlrd.open_workbook(file_title)
data = wb.sheet_by_index(0) # This code block opens a sheet. 
b = Workbook()
w = b.add_sheet("sheet")

for i in range(1, data.nrows): # This block creates and outputs a sheet (formatted.xls)
    w.write(i, 0, data.cell_value(i,0))
    w.write(i, 1, data.cell_value(i,1))
    w.write(i, 2, data.cell_value(i,2).upper())
    w.write(i, 3, data.cell_value(i,3).upper())
    w.write(i, 4, data.cell_value(i,4).upper())
    w.write(i, 5, standardize([data.cell_value(i, 5), data.cell_value(i, 6), data.cell_value(i, 7), str(int(data.cell_value(i, 8)))]))
    w.write(i, 6, data.cell_value(i,6).upper())
    w.write(i, 7, data.cell_value(i,7).upper())
    w.write(i, 8, data.cell_value(i,8))
    # standardize([data.cell_value(i, 5), data.cell_value(i, 6), data.cell_value(i, 7), str(data.cell_value(i, 8))])

b.save("formatted.xls")
print("---")
print(len(difficult_addresses))
print(difficult_addresses)
