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
    """QRCODE ATTENDANCE RECORDING AUTOMATION PROGRAM
Programmed and maintained by
Adrian Luke Labasan | G11-Oxygen | CONTACT: zionexodus7@protonmail.com\n
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     ??? INSTRUCTIONS ???
* Place the excel file at the same directory where the program (qrcode_attendance.py) is located.
* If first used, the excel file should be completely blank i.e. no data added.
* Press Ctrl+C to end the program.
* If you would like to change the layout of the excel file, please notify the programmer.

     !!!   WARNING    !!!
* Don't manipulate the excel file when the program is running!
* Before running the program, ensure that:
   (1) The excel file is closed.
   (2) The excel file is not manipulated/changed in any way beforehand except by the program.

     ***    NOTE      ***
* The program only records the initial time a qrcode is scanned during the day.
* The order of the names are based on the time the qrcodes are first scanned. If you would like to sort the names by their section, you can do so in Excel: Select data range (include the times recorded) -> Data tab -> SORT

If you want to suggest additional functionalities or modifications to the program, feel free to contact the programmer :)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n"""
)

def getExcelFiles():
    excelFiles = []
    for file in os.listdir('.'):
        if file.endswith('.xlsx'):
            excelFiles.append(file)
    return excelFiles

while True:
    excelFiles = getExcelFiles()

    noInitialFile = False

    if len(excelFiles) == 0:
        print('No xlsx files found in the directory.')
        print('Please enter a filename for the excel file to be created.')
        print('Press enter if you would like the default name [GARSHS_ATTENDANCE.xlsx]')
        excelFilename = pyinputplus.inputStr(prompt='> ')
        if not excelFilename.endswith('.xlsx'):
            excelFilename += '.xlsx'
        noInitialFile = True
        break
    elif len(excelFiles) == 1:
        excelFilename = excelFiles[0]
        print(f'Found {excelFilename}')
        print('Exit the program (Ctrl+C) if this is not the excel file you want.')
        break
    elif len(excelFiles) == 2 and (excelFiles[0].startswith('~$') or excelFiles[1].startswith('~$')):
        print('!!! The excel file is open in another window !!!')
        print('Please close the excel file then press enter to continue.')
        continueProgram = input('> ')
        continue
    else:
        exelFileIsOpen = False
        for file in os.listdir('.'):
            if exelFileIsOpen:
                break
            if file.startswith('~$') and file.endswith('.xlsx'):
                print('!!! An excel file is open in another window !!!')
                print('Please close the excel file/s then press enter to continue.')
                continueProgram = input('> ')
                for file in os.listdir('.'):
                    if file.startswith('~$') and file.endswith('.xlsx'):
                        exelFileIsOpen = True
                        break
        if exelFileIsOpen:
            continue
        print('Multiple xlsx files found in the directory.')
        excelFilename = pyinputplus.inputMenu(getExcelFiles(), numbered=True)
        break

print("\nOpening the webcam... please wait\n")
cap = cv2.VideoCapture(0)
cap.set(3, 640)
cap.set(4, 480)

attendance = {}


def getData(ws):
    rows = list(ws.rows)
    rowsDict = {}
    for i, row in enumerate(rows):
        rowsDict.setdefault(i + 1, [])
        for cell in row:
            rowsDict[i + 1].append(cell.value)

    columns = list(ws.columns)
    columnsDict = {}
    for i, column in enumerate(columns):
        columnsDict.setdefault(i + 1, [])
        for cell in column:
            columnsDict[i + 1].append(cell.value)
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

        cv2.imshow("Please show your QR code to the webcam.", img)
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
            print("\n!!!  The excel file is open on another window  !!!")
            print("Close the excel file to let the program manipulate it.")
            print("Press enter if done closing the window.")
            continueProgram = input('> ')
            continue
        else:
            for date in attendance:
                columnIndex = 0
                for column_index in columnsDict:
                    if date == columnsDict[column_index][1]:
                        columnIndex = column_index

                for name in attendance[date]:
                    rowIndex = 0
                    for row_index in rowsDict:
                        if name == rowsDict[row_index][0]:
                            rowIndex = row_index
                    ws.cell(row=rowIndex, column=columnIndex).value = attendance[date][name]

            wb.save(excelFilename)
            break

print(f"\nExcel file {'created' if noInitialFile else 'updated'}.\nThank you. Have a good day :)")
