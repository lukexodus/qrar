import logging
import datetime
import shelve

import cv2
import numpy as np
from pyzbar.pyzbar import decode
import openpyxl

# logging.basicConfig(filename='log.txt', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logging.disable(logging.DEBUG)

with open("excel_filename.txt") as file:
    excelFilename = file.read().strip()
    logging.debug(f"'{excelFilename}'")

print(
    """Program started.\n
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     ??? INSTRUCTIONS ???
print("Please don't manipulate the excel file while the program is running!
Press Ctrl+C to end the program.
If you would like to change the layout of the excel file, please notify the programmer: Adrian Luke Labasan (G11-Oxygen)
If you would like to change the filename of the excel file, please change it in the excel_filename.txt file before renaming the excel file.

     !!!     NOTE     !!!
Before running the program, ensure that:
/   The excel file is closed.")
/   The excel file is not manipulated/changed in any way beforehand.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n"""
)

print("Opening the webcam... please wait\n")
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
        wb = openpyxl.load_workbook(excelFilename)
        ws = wb.active

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
print("\nExcel file updated.\nProgram ended.")
