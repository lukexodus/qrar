import logging
import datetime

import cv2
import numpy as np
from pyzbar.pyzbar import decode
import openpyxl

# logging.basicConfig(filename='log.txt', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logging.disable(logging.INFO)

cap = cv2.VideoCapture(0)
cap.set(3, 640)
cap.set(4, 480)

wb = openpyxl.load_workbook('GARSHS_ATTENDANCE.xlsx')
ws = wb.active

attendance = {}
logging.info(f'TEMP attendance -> {attendance}\n')


def getData(ws):
    rows = list(ws.rows)
    rowsDict = {}
    for i, row in enumerate(rows):
        rowsDict.setdefault(i+1, [])
        for cell in row:
            rowsDict[i+1].append(cell.value)
    logging.debug(f'rowsDict -> {rowsDict}\n')

    columns = list(ws.columns)
    columnsDict = {}
    for i, column in enumerate(columns):
        columnsDict.setdefault(i+1, [])
        for cell in column:
            columnsDict[i+1].append(cell.value)
    logging.debug(f'columnsDict -> {columnsDict}\n')
    return rowsDict, columnsDict


try:
    while True:
        success, img = cap.read()
        for qrcode in decode(img):
            name = qrcode.data.decode('utf-8')

            currentTime = datetime.datetime.now()
            date = currentTime.strftime('%a %Y/%m/%d')
            time = currentTime.strftime('%H:%M:%S')
            """ logging.info(f'TEMP attendance[date] = {attendance[date]}')
            logging.info(f'attendance[date] = {attendance[date]}')
            logging.info(f'name = {name}') """
            attendance.setdefault(date, {})
            if name not in attendance[date]:
                print(name, date, time)
            attendance[date][name] = time

            pts = np.array([qrcode.polygon], np.int32)
            pts = pts.reshape((-1, 1, 2))
            cv2.polylines(img, [pts], True, (255, 0, 255), 5)
            pts2 = qrcode.rect
            cv2.putText(img, name, (pts2[0], pts2[1]), cv2.FONT_HERSHEY_COMPLEX, 0.9, (255, 0, 255), 2)

        cv2.imshow("GARSHS Attendance", img)
        cv2.waitKey(1)
except KeyboardInterrupt:
    logging.debug(f'attendance -> {attendance}\n')

    rowsDict, columnsDict = getData(ws)

    for date in attendance.keys():
        if date not in rowsDict[2]:
            ws.cell(row=2, column=ws.max_column+1).value = date
    for date in attendance:
        for name in attendance[date].keys():
            if name not in columnsDict[1]:
                ws.cell(row=ws.max_row+1, column=1).value = name

    wb.save('GARSHS_ATTENDANCE.xlsx')
    wb = openpyxl.load_workbook('GARSHS_ATTENDANCE.xlsx')
    ws = wb.active
    rowsDict, columnsDict = getData(ws)

    for date in attendance:
        logging.debug(f'date = {date}')
        columnIndex = 0
        for column_index in columnsDict:
            logging.debug(f'column_index = {column_index}')
            logging.debug(f'columnsDict[column_index][1] = {columnsDict[column_index][1]}')
            if date == columnsDict[column_index][1]:
                columnIndex = column_index
                logging.debug(f'located: columnIndex = {columnIndex}')
        logging.debug(f'after: columnIndex = {columnIndex}')

        for name in attendance[date]:
            logging.debug(f'name = {name}')
            rowIndex = 0
            for row_index in rowsDict:
                logging.debug(f'row_index = {row_index}')
                logging.debug(f'rowsDict[row_index][0] = {rowsDict[row_index][0]}')
                if name == rowsDict[row_index][0]:
                    rowIndex = row_index
                    logging.debug(f'located: rowIndex = {rowIndex}')
            logging.debug(f'after: rowIndex = {rowIndex}')
            logging.debug(f'cell: {columnIndex}, {rowIndex}')
            logging.debug(f'cell: {attendance[date][name]}')
            ws.cell(row=rowIndex, column=columnIndex).value = attendance[date][name]

    wb.save('GARSHS_ATTENDANCE.xlsx')