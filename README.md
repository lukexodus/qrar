qrar
======

A python program which automates attendance recording by transfering decoded data from QR codes to an excel file and a Google sheet, with email notification feature.

Installation
------------

To install with pip, run:

    pip install qrar

Quickstart Guide
----------------

# Instructions #
1) Place the excel file at the same directory as the program (qrar.py).
2) If first used, the excel file should be completely blank i.e. no data added.
3) Press Ctrl+C to end the program.
4) If you would like to change the layout of the excel file, please notify the programmer.

# Note #
1) The program only records the initial time a qrcode is scanned during the day.
2) The order of the names are based on the time the qrcodes are first scanned. If you would like to sort the names alphabetically (ex. by sections), you can do so in Excel: [(Select data range by row) -> 'Data' tab -> 'SORT']. You can also do this in Google Sheets: [(Select data range by row) -> 'Data' tab -> 'Sort range'] but you don't have to since the program will automatically updates the changes in the local database to the online databse
3) The program reduces the row and column counts in the online spreadsheet to make the read/write process faster. The program will automatically increase the row and column counts as the data grows.
4) If you use the email notification function:
   a) You can edit the EDIT_EMAIL_NOTIF_MESSAGE.txt to change the subject and body of the message of the notification to be sent.
   b) You can only use the email notification function after you had added the emails of the members. However, you can only add the members' email and phone number after the xlsx file is initialized (if already created by the user or by the program).
   c) The program only sends email to those whose email is added to the xlsx file.

# Warning #
1) Don't manipulate the excel file while the program is running.
2) Before running the program, ensure that:
   a) The excel file is closed.
   b) The excel file is not manipulated/changed in any way beforehand except by the program.
3) DO NOT DELETE the backup.* files. These files are used in the backup system and in the sync program. Delete the files only if you want to use the program for another set of data i.e. new set of students/members, new schoolyear, etc. (It won't affect the xlsx file.)
4) The googleSheetsConfig.* files contain your configuration for the online database (Google Sheets). Delete these files if you want to reset them.
5) The emailNotifConfig.* files contain your configuration for the email notification function (Gmail). Delete these files if you want to reset them.
6) The settings.* files contain the settings of the behavior of the program. Do not delete these files also.
7) DO NOT DELETE the token files. They contain your sign-in configuration for Google Sheets and Gmail.
8) If you want to delete the spreadsheet in Google Sheets (ex. if you want to reupload it), you have to delete it also in the trash folder in your google drive. If you don't, the program can still detect it and so uses it as its online database.

Contribute
----------

If you'd like to contribute to qrar, check out https://github.com/shiideyuuki/qrar
