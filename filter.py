import sys
import win32com.client as win32

"""
The script aimed to update slicer filtering according to BOM items.

Version 0.1 beta
"""

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

data = ["Vegetables"]

# open appropriate excel document
try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = openWorkbook(excel, 'script.xlsx')
    ws = wb.Worksheets('Ortho')
    sl = wb.SlicerCaches("Slicer_BOM")

    # working example of visible slicer item list 
    sl.VisibleSlicerItemsList = ["[Medical Case].[BOM].&[SD900.104 (Qty:1)]"]

except Exception as e:
   print(e)

finally:
    # RELEASES RESOURCES
    #wb.Close(SaveChanges=1)
    ws = None
    wb = None
    excel = None



