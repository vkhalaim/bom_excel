# -*- coding: utf-8 -*-
import os
import win32com.client as win32
from tkinter import filedialog as fd
from pathlib import Path

FILE_NAME = ''
WORK_BOOK = None

data_folder = Path("./txt")
# bom elements which should be turned on after filtering - guides/models
GUIDES = ["SD900.101", "SD900.102", "SD900.104", "SD900.105", "SD900.106",
          "SD900.107", "SD900.108", "SD900.109", "SD900.110", "SD900.111",
          "SD900.001", "SD900.003", "SD900.004", "SD900.006", "SD900.008",
          "SD900.009", "SD900.010", "SD900.011", "SD900.051", "SD900.054",
          "SD900.056", "SD980.001", "SD980.002", "SD980.005", "SD980.006",
          "SD980.009", "SD980.120", "OBL031-F", "OBL032-F", "OBL033-F",
          "OBL034-F", "OBL035-F", "OBL171-F", "OBL172-F", "OBL173-F",  "OBL174-F",
          "OBL175-F", "OBL041-F"]
MODELS = ['SD900.201', 'SD900.202', 'SD900.203', 'SD900.204', 'SD900.205',
          'SD900.206', 'SD900.207', 'SD900.208', 'SD900.209', 'SD900.212',
          'SD900.213', 'SD900.214', 'SD900.215', 'SD900.216', 'SD900.231',
          'SD900.236', 'SD900.237', 'SD900.238', 'SD900.239', 'SD900.261',
          'SD900.262', 'SD900.263', 'SD900.264', 'SD900.265', 'SD900.266',
          'SD900.267', 'SD900.268', 'SD900.269', 'SD900.270', 'SD900.301',
          'SD900.302', 'SD900.303', 'SD900.304', 'SD900.305', 'SD900.306',
          'SD900.307', 'SD900.308', 'SD900.309', 'SD900.312', 'SD900.331',
          'SD900.336', 'SD900.337', 'SD900.338', 'SD900.339', 'SD900.361',
          'SD900.362', 'SD900.363', 'SD900.364', 'SD900.365', 'SD900.366',
          'SD900.367', 'SD900.368', 'SD900.369', 'SD900.370', 'SD900.381',
          'SD900.383', 'SD900.384', 'SD900.385', 'OBL011-F', 'OBL012-F',
          'OBL013-F', 'OBL014-F', 'OBL015-F', 'OBL016-F', 'OBL017-F', 'OBL021-F',
          'OBL022-F', 'OBL023-F', 'OBL024-F', 'OBL025-F', 'OBL026-F', 'OBL027-F']


def not_in_guides(line):
    for guide in GUIDES:
        if guide in line:
            return False

    return True


def get_logo_path():
    return Path('./img') / 'materialise-logo.png'


def open_file():
    file_name = fd.askopenfile()
    global FILE_NAME
    global WORK_BOOK
    FILE_NAME = file_name.name
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    try:
        WORK_BOOK = excel.Workbooks(FILE_NAME)
    except Exception as e:
        try:
            WORK_BOOK = excel.Workbooks.Open(FILE_NAME)
        except Exception as e:
            print(e)
            WORK_BOOK = None


def filtering(SLICER_NAME):
    # filling array with already filtered BOM data
    data = []

    with open(data_folder / 'bom_filtered.txt', 'r') as my_file:
        for line in my_file:
            data.append(line.rstrip('\n'))

    # open appropriate excel document
    try:
        wb = WORK_BOOK
        sl = wb.SlicerCaches(SLICER_NAME)
        # select only needed data in slicer
        sl.VisibleSlicerItemsList = data

    except Exception as e:
        print(e)

    finally:
        # RELEASES RESOURCES
        wb = None


def prepare(SLICER_NAME, MODE):
    allSlicerElements = ()

    try:
        wb = WORK_BOOK
        sl = wb.SlicerCaches(SLICER_NAME)
        # select all elements from slicer
        allSlicerElements = sl.VisibleSlicerItemsList

    except Exception as e:
        print(e)

    # remove all text from file before writing new info
    with open(data_folder / 'bom_items.txt', 'w') as f:
        print(allSlicerElements)
        f.truncate(0)

        for elem in allSlicerElements:
            f.write(elem + '\n')

    bom_array = []

    if MODE == 'Guides':
        with open(data_folder / 'bom_items.txt', 'r') as my_file:
            for line in my_file:
                for elem in GUIDES:
                    if elem in line:
                        bom_array.append(line)
                        break

        with open(data_folder / 'bom_filtered.txt', 'w') as f:
            f.truncate(0)

            for item in bom_array:
                f.write(item)

    if MODE == 'Models':
        with open(data_folder / 'bom_items.txt', 'r') as my_file:
            for line in my_file:
                for model in MODELS:
                    if model in line and not_in_guides(line):
                        bom_array.append(line)
                        break

        with open(data_folder / 'bom_filtered.txt', 'w') as f:
            f.truncate(0)

            for item in bom_array:
                f.write(item)
