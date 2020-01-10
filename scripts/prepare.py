import sys
import win32com.client as win32
from pathlib import Path

from scripts.filter import openWorkbook

''' Filtering needed BOM items from file '''

bom_needed = ["SD900.101", "SD900.102", "SD900.104", "SD900.105", "SD900.106", "SD900.107", "SD900.108", "SD900.109", "SD900.110", "SD900.111", "SD900.001", 
    "SD900.003", "SD900.004", "SD900.006", "SD900.008", "SD900.009", "SD900.010", "SD900.011", "SD900.051", "SD900.054", "SD900.056", "SD980.001", "SD980.002", 
    "SD980.005", "SD980.006", "SD980.009", "SD980.120"
]
data_folder = Path("./txt")

def prepare(FILE_NAME, SLICER_NAME):
    allSlicerElements = ()
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = openWorkbook(excel, Path(FILE_NAME))
        sl = wb.SlicerCaches(SLICER_NAME)

        allSlicerElements = sl.VisibleSlicerItemsList # select all elements from slicer
        
    except Exception as e:
        print(e)

    # remove all text from file before writing new info
    with open(data_folder / 'bom_items.txt', 'w') as f:
        print(allSlicerElements)
        f.truncate(0)

        for elem in allSlicerElements:
            f.write(elem + '\n')

        
    # fill in array with needed elements
    bom_array = []

    with open(data_folder / 'bom_items.txt', 'r') as my_file:
        for line in my_file:
            for elem in bom_needed:
                if elem in line:
                    bom_array.append(line)
                    break


    #erase bom filtered elemets before write
    with open(data_folder / 'bom_filtered.txt', 'w') as f:
        f.truncate(0)

        for item in bom_array:
            f.write(item)
