import sys

from win32com.client import Dispatch

"""
The script aimed to update slicer filtering according to BOM items.

"""
data = ["Fruit", "Vegetables"]

xl = Dispatch("Excel.Application")
wb = xl.Workbooks.Open("example.xlsx")
sl = wb.SlicerCaches("Slicer_Category")

for it in sl.SlicerItems:
    print(it.name)
    if it.name in data:
        it.Selected = True
    else:
        it.Selected = False
