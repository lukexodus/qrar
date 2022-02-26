import logging
import datetime
import shelve
import os

import cv2
import numpy as np
from pyzbar.pyzbar import decode
import openpyxl
from openpyxl.styles import Font
import pyinputplus

# logging.basicConfig(filename='log.txt', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logging.disable(logging.DEBUG)

print(
    """Program started.\n
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     ??? INSTRUCTIONS ???
* Place the excel file at the same directory where the program (qrcode_attendance.py) is located.
* If first used, the excel file should be completely blank i.e. not changed in any way.
* Press Ctrl+C to end the program.
* If you would like to change the layout/template of the excel file, please notify first the programmer: Adrian Luke Labasan (G11-Oxygen) <zionexodus7@protonmail.com>

     !!!     NOTE     !!!
* Please don't manipulate the excel file while the program is running!

Before running the program, ensure that:
/   The excel file is closed.
/   The excel file is not manipulated/changed in any way beforehand except by the program.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n"""
)

excelFiles = []
for file in os.listdir('.'):
    if file.endswith('.xlsx'):
        excelFiles.append(file)

noInitialFile = False

if len(excelFiles) == 0:
    print('No xlsx files found in the directory.')
    print('Please enter a filename for the excel file to be created.')
    print('Press enter if you would like the default name [GARSHS_ATTENDANCE.xlsx]')
    excelFilename = pyinputplus.inputStr(prompt='> ')
    if not excelFilename.endswith('.xlsx'):
        excelFilename += '.xlsx'
    noInitialFile = True
elif len(excelFiles) == 1:
    excelFilename = excelFiles[0]
    print(f'Found {excelFilename}')
    print('Exit the program (Ctrl+C) if this is not the excel file you want.')
else:
    print('Multiple xlsx files found in the directory.')
    excelFilename = pyinputplus.inputMenu(excelFiles, numbered=True)

print("\nOpening the webcam... please wait\n")
cap = cv2.VideoCapture(0)
cap.set(3, 640)
cap.set(4, 480)

attendance = {}
logging.debug(f"TEMP attendance -> {attendance}\n")


def getData(ws):
    rows = list(ws.rows)
    rowsDict = {}
    for i, row in enumerate(rows):
        rowsDict.setdefault(i + 1, [])
        for cell in row:
            rowsDict[i + 1].append(cell.value)
    logging.debug(f"rowsDict -> {rowsDict}\n")

    columns = list(ws.columns)
    columnsDict = {}
    for i, column in enumerate(columns):
        columnsDict.setdefault(i + 1, [])
        for cell in column:
            columnsDict[i + 1].append(cell.value)
    logging.debug(f"columnsDict -> {columnsDict}\n")
    return rowsDict, columnsDict


try:
    print("Webcam is running. You can now show your QR Code to the webcam.")
    print("--------------------ATTENDANCE_LOG--------------------")
    while True:
        success, img = cap.read()
        for qrcode in decode(img):
            name = qrcode.data.decode("utf-8")

            currentTime = datetime.datetime.now()
            date = currentTime.strftime("%a %Y/%m/%d")
            time = currentTime.strftime("%H:%M:%S")
            attendance.setdefault(date, {})
            if name not in attendance[date]:
                print(date, time, name)
            attendance[date][name] = time

            pts = np.array([qrcode.polygon], np.int32)
            pts = pts.reshape((-1, 1, 2))
            cv2.polylines(img, [pts], True, (255, 0, 255), 5)
            pts2 = qrcode.rect
            cv2.putText(
                img,
                name,
                (pts2[0], pts2[1]),
                cv2.FONT_HERSHEY_COMPLEX,
                0.9,
                (255, 0, 255),
                2,
            )

        cv2.imshow("GARSHS Attendance", img)
        cv2.waitKey(1)
except KeyboardInterrupt:
    while True:
        if noInitialFile:
            wb = openpyxl.Workbook()
            ws = wb.active
        else:
            wb = openpyxl.load_workbook(excelFilename)
            ws = wb.active

        if ws['A1'].value != '':
            ws['A1'].value = 'Attendance'
            ws['A1'].font = Font(size=20)
        if ws['A2'].value != '':
            ws['A2'].value = 'Name'
            ws['A2'].font = Font(size=16)

        logging.debug(f"attendance -> {attendance}\n")

        rowsDict, columnsDict = getData(ws)

        for date in attendance.keys():
            if date not in rowsDict[2]:
                ws.cell(row=2, column=ws.max_column + 1).value = date
        for date in attendance:
            for name in attendance[date].keys():
                if name not in columnsDict[1]:
                    ws.cell(row=ws.max_row + 1, column=1).value = name
        try:
            wb.save(excelFilename)
            wb = openpyxl.load_workbook(excelFilename)
            ws = wb.active
            rowsDict, columnsDict = getData(ws)
        except PermissionError:
            """permissionError_shelfFIle = shelve.open('PermissionError_backup')
            for date in attendance:
                permissionError_shelfFIle[date] = attendance[date]
            permissionError_shelfFIle.close()"""
            print("\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
            print("The excel file is open on another window. Close the excel file to let the program manipulate it.")
            print("Press enter if done closing the window.")
            print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
            continueProgram = input()
            continue
        else:
            for date in attendance:
                logging.debug(f"date = {date}")
                columnIndex = 0
                for column_index in columnsDict:
                    logging.debug(f"column_index = {column_index}")
                    logging.debug(f"columnsDict[column_index][1] = {columnsDict[column_index][1]}")
                    if date == columnsDict[column_index][1]:
                        columnIndex = column_index
                        logging.debug(f"located: columnIndex = {columnIndex}")
                logging.debug(f"after: columnIndex = {columnIndex}")

                for name in attendance[date]:
                    logging.debug(f"name = {name}")
                    rowIndex = 0
                    for row_index in rowsDict:
                        logging.debug(f"row_index = {row_index}")
                        logging.debug(f"rowsDict[row_index][0] = {rowsDict[row_index][0]}")
                        if name == rowsDict[row_index][0]:
                            rowIndex = row_index
                            logging.debug(f"located: rowIndex = {rowIndex}")
                    logging.debug(f"after: rowIndex = {rowIndex}")
                    logging.debug(f"cell: {columnIndex}, {rowIndex}")
                    logging.debug(f"cell: {attendance[date][name]}")
                    ws.cell(row=rowIndex, column=columnIndex).value = attendance[date][name]

            wb.save(excelFilename)
            break

print(f"\nExcel file {'created' if noInitialFile else 'updated'}.\nProgram ended.")
