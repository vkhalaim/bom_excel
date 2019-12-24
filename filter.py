import sys
import win32com.client as win32

"""
The script aimed to update slicer filtering according to BOM items.

Version 1.0
"""
FILE_NAME = "script.xlsx"
SLICER_NAME = "Slicer_BOM"

def openWorkbook(xlapp, xlfile):
    try:        
        xlwb = xlapp.Workbooks(xlfile)            
    except Exception as e:
        try:
            xlwb = xlapp.Workbooks.Open(xlfile)
        except Exception as e:
            print(e)
            xlwb = None                    
    return(xlwb)

# filling array with already filtered BOM data
data = []

with open('bom_filtered.txt') as my_file:
    for line in my_file:
        data.append(line.rstrip('\n'))

# open appropriate excel document
try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = openWorkbook(excel, FILE_NAME)
    sl = wb.SlicerCaches(SLICER_NAME)
    sl.VisibleSlicerItemsList = data # select only needed data in slicer
    

except Exception as e:
   print(e)

finally:
    # RELEASES RESOURCES
    ws = None
    wb = None
    excel = None



