import logging
import datetime
import shelve
import os
from tkinter import FALSE
import time
import traceback
from contextlib import suppress
import threading
from typing import List, Dict

import cv2
import numpy as np
from pyzbar.pyzbar import decode
import openpyxl
from openpyxl.styles import Font
import pyinputplus
import ezsheets
from ezsheets import getColumnLetterOf, getColumnNumberOf
from ezsheets import EZSheetsException
import smtplib
import httplib2
import pyinputplus as pyip
from google.auth.exceptions import TransportError
from ssl import SSLEOFError

# logging.basicConfig(filename='log.txt', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
logging.disable(logging.DEBUG)

print(
    """QRCODE ATTENDANCE RECORDING AUTOMATION PROGRAM
Programmed and maintained by
Adrian Luke Labasan | G11-Oxygen | CONTACT: zionexodus7@protonmail.com ~ 09157694749\n
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
* The order of the names are based on the time the qrcodes are first scanned. If you would like to sort the names by their section (alphabetically), you can do so in Excel: (Select data range by row) -> 'Data' tab -> 'SORT'. In Google Sheets: (Select data range by row) -> 'Data' tab -> 'Sort range'.

If you want to suggest additional functionalities or modifications to the program, feel free to contact the programmer :)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n"""
)


# Local database interface
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


# Online database interface
def countdown(seconds=5):
    print(f'Retrying in {seconds} ', end='', flush=True)
    time.sleep(1)
    count = seconds - 1
    while count > 0:
        print(f'{count} ', end='', flush=True)
        count = count - 1
        time.sleep(1)


def silentCountdown(seconds=5):
    for second in range(seconds):
        time.sleep(1)

def credentialsSheetsInDirectory():
    for filename in os.listdir('.'):
        if filename == 'credentials-sheets.json':
            return True
    return False


useOnlineDatabase = True
if not credentialsSheetsInDirectory():
    exitSetup = False
    response1 = pyip.inputYesNo('Do you want to use an online database (Google Sheets)?\n> ')
    while not exitSetup:
        if response1 == 'no':
            useOnlineDatabase = False
            break
        else:
            print("""Please follow these steps to be able to use Google Sheets:
1) Enable Google Sheets and Google Drive APIs through these links:
    * https://console.developers.google.com/apis/library/sheets.googleapis.com/
    * https://console.developers.google.com/apis/library/drive.googleapis.com/
2) Create and download credential file in this link https://console.cloud.google.com/apis/credentials/oauthclient
    - You may have to configure the consent screen to create the credential file
3) Rename the downloaded file (json) to 'credentials-sheets.json' and place it in the same folder as the program
    - After you continue the program, you would have to sign in and authorize the app (twice)

NOTE: Always connect to the internet if you plan to use an online database.""")
            print('Press enter if done.')
            input('> ')
            exitSetup = credentialsSheetsInDirectory()

# Checks if the xlsx file is already uploaded to Google Sheets
# and if it is, gets its spreadsheetId
def spreadsheetExists():
    spreadsheets = ezsheets.listSpreadsheets()
    for i, key in enumerate(spreadsheets):
        if spreadsheets[key] == excelFilename[:-5]:
            return True, key
    return False, None


if useOnlineDatabase:
    try:
        with shelve.open('backup') as backup:
            uploadedToGoogleSheets = backup['uploadedToGoogleSheets']
    except KeyError:
        print('\nDetecting if the xlsx file is already uploaded to Google Sheets...')
        uploadedToGoogleSheets = False
        while True:
            try:
                uploadedToGoogleSheets, spreadsheetId = spreadsheetExists()
                if uploadedToGoogleSheets:
                    with shelve.open('backup') as backup:
                        backup['uploadedToGoogleSheets'] = True
                        backup['spreadsheetId'] = spreadsheetId
                else:
                    with shelve.open('backup') as backup:
                        backup['uploadedToGoogleSheets'] = False
                break
            except (ConnectionAbortedError, httplib2.ServerNotFoundError, httplib2.error.ServerNotFoundError, TransportError, SSLEOFError) as error:
                with open('log.txt', 'a') as log:
                    log.write(traceback.format_exc())
                print('Connection error. Please check your internet connection.')
                countdown()
                continue


class ExtendedSheet:
    def __init__(self, ss, title):
        self.sheet = ss[title]

    
    def dateInColumnHeaders(self, date):
        row2 = self.sheet.getRow(2)
        for cell in row2:
            if cell == date:
                return True
        return False
    

    def nameInRowHeaders(self, name):
        col1 = self.sheet.getColumn(1)
        for cell in col1:
            if cell == name:
                return True
        return False


    def findFirstBlankCol(self):
        row2 = self.sheet.getRow(2)
        colNum = 1
        while row2[colNum - 1] != '':
            colNum += 1
        return colNum
    

    def findFirstBlankRow(self):
        col1 = self.sheet.getColumn(1)
        rowNum = 1
        while col1[rowNum - 1] != '':
            rowNum += 1
        return rowNum



if useOnlineDatabase and uploadedToGoogleSheets:
    while True:
        with shelve.open('backup') as backup:
            spreadsheetId = backup['spreadsheetId']
        try:
            print('\nConnecting to the online database...')
            ss = ezsheets.Spreadsheet(spreadsheetId)
            sheet1 = ExtendedSheet(ss, 'Sheet')
            break
        except (httplib2.ServerNotFoundError, httplib2.error.ServerNotFoundError, TransportError, SSLEOFError) as error:
            with open('log.txt', 'a') as log:
                log.write(traceback.format_exc())
            print('Failed to connect to Google Sheets. Please check your internet connection.')
            countdown()
            continue


# Retrives the backup 'attendance' dictionary, otherwise initializes it
with shelve.open('backup') as backup:
    try:
        attendance = backup['attendance']
    except KeyError:
        attendance = {}


# Opens the webcam to scan QR codes
print("\nOpening the webcam... please wait\n")
cap = cv2.VideoCapture(0)
cap.set(3, 640)
cap.set(4, 480)

try:
    print("Webcam is running. You can now show your QR Code to the webcam.")
    print("------------------------ATTENDANCE_LOG-------------------------")

    dateCheck = ''  # 'date' cache, used to check 'date' change

    while True:
        success, img = cap.read()
        for qrcode in decode(img):
            # Appends date information to the decoded data and stores
            # it in the dictionary 'attendance'
            name = qrcode.data.decode("utf-8")
            currentTime = datetime.datetime.now()
            date = currentTime.strftime("%a %Y/%m/%d")
            clockTime = currentTime.strftime("%H:%M:%S")
            attendance.setdefault(date, {})
            if name not in attendance[date]:
                print(date, clockTime, name)
                attendance[date][name] = clockTime

                def writeData(sheet1, name, date, dateCheck, clockTime):
                    # Adds the date in the headers if the date changes while the 
                    # program is running i.e. during midnight
                    while True:
                        try:
                            sheet1.sheet.refresh()
                            if date != dateCheck:
                                dateCheck = date
                                if not sheet1.dateInColumnHeaders(dateCheck):
                                    colNum = sheet1.findFirstBlankCol()
                                    sheet1.sheet.columnCount = colNum + 2
                                    sheet1.sheet[f'{getColumnLetterOf(colNum)}2'] = date

                            if not sheet1.nameInRowHeaders(name):
                                rowNum = sheet1.findFirstBlankRow()
                                sheet1.sheet[f'A{rowNum}'] = name
                                sheet1.sheet.rowCount = rowNum + 2

                            row2 = sheet1.sheet.getRow(2)
                            col1 = sheet1.sheet.getColumn(1)
                            colIndex = row2.index(date) + 1
                            rowIndex = col1.index(name) + 1
                            sheet1.sheet[f'{getColumnLetterOf(colIndex)}{rowIndex}'] = clockTime
                            break
                        except (ConnectionResetError, ConnectionAbortedError, httplib2.ServerNotFoundError, httplib2.error.ServerNotFoundError, TransportError, SSLEOFError) as error:
                            print('Connection error. Please check your internet connection.')
                            silentCountdown()
                            continue
            
                # Real-time upload of data to the online database
                if useOnlineDatabase and uploadedToGoogleSheets:
                    threading.Thread(target=writeData, args=[sheet1, name, date, dateCheck, clockTime]).start()


            # Configure the design and information showed in the camera
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
    # Backups the gathered data
    with shelve.open('backup') as backup:
        backup['attendance'] = attendance
        logging.debug(backup['attendance'])
    

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


    while True:
        # Initializes the local database if not already initialized
        if noInitialFile:
            wb = openpyxl.Workbook()
            ws = wb.active
        else:
            wb = openpyxl.load_workbook(excelFilename)
            ws = wb.active
        ws.freeze_panes = 'B3'

        if ws['A1'].value != '':
            ws['A1'].value = 'Attendance'
            ws['A1'].font = Font(size=20)
        if ws['A2'].value != '':
            ws['A2'].value = 'Name'
            ws['A2'].font = Font(size=16)
        if ws['B2'].value != '':
            ws['B2'].value = 'Email'
            ws['B2'].font = Font(size=14)
        if ws['C2'].value != '':
            ws['C2'].value = 'Phone #'
            ws['C2'].font = Font(size=14)


        # Adds the necessary locators (row: name & column: date)
        rowsDict, columnsDict = getData(ws)
        for date in attendance.keys():
            if date not in rowsDict[2]:
                ws.cell(row=2, column=ws.max_column + 1).value = date
        for date in attendance:
            for name in attendance[date].keys():
                if name not in columnsDict[1]:
                    ws.cell(row=ws.max_row + 1, column=1).value = name

        # Writes the gathered data to the local database if the xlsx file
        # is not currently open, otherwise restarts the initialization
        try:
            wb.save(excelFilename)
            wb = openpyxl.load_workbook(excelFilename)
            ws = wb.active
            rowsDict, columnsDict = getData(ws)
        except PermissionError:
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
            print(f"\nExcel file {'created' if noInitialFile else 'updated'}.")
            time.sleep(1)


            # Uploads the xlsx file to the online database if not already uploaded
            # (if configured to use oneline database)
            if useOnlineDatabase and not uploadedToGoogleSheets:
                while True:
                    try:
                        print('Uploading excel file to Google Sheets...')
                        ss = ezsheets.upload(excelFilename)
                    except (httplib2.ServerNotFoundError, httplib2.error.ServerNotFoundError) as error:
                        with open('log.txt', 'a') as log:
                            log.write(traceback.format_exc())
                        print('Upload to Google Sheets failed. Please check your internet connection.')
                        countdown()
                        continue
                    except (TransportError, SSLEOFError) as error:
                        while True:
                            try:
                                uploadedToGoogleSheets, spreadsheetId = spreadsheetExists()
                                if uploadedToGoogleSheets:
                                    with shelve.open('backup') as backup:
                                        backup['uploadedToGoogleSheets'] = True
                                else:
                                    with shelve.open('backup') as backup:
                                        backup['uploadedToGoogleSheets'] = False
                                break
                            except (ConnectionAbortedError, httplib2.ServerNotFoundError, httplib2.error.ServerNotFoundError, TransportError, SSLEOFError) as error:
                                with open('log.txt', 'a') as log:
                                    log.write(traceback.format_exc())
                                print('Update to online database failed. Please check your internet connection.')
                                countdown(8)
                                continue
                    else:
                        with shelve.open('backup') as backup:
                            backup['uploadedToGoogleSheets'] = True
                            backup['spreadsheetId'] = ss.spreadsheetId
                        try:
                            with shelve.open('backup') as backup:
                                setFrozenPanes = backup['setFrozenPanes']
                        except KeyError:
                            ss['Sheet'].frozenRowCount = 2
                            ss['Sheet'].frozenColumnCount = 1
                            with shelve.open('backup') as backup:
                                backup['setFrozenPanes'] = True
                        print('Upload success.')
                        break
            break

print("\nThank you. Have a good day :)")
