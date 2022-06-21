# +++--MISMATCH VERSION ISSUE--+++
# BECAUSE WE HAVE MACHINE AND PYTHON 64-bit, BUT MICROSOFT ACCESS USE 32-bit(MOST OF THEM RUN IN THIS WAY),
# WE NEED TO INSTALL FILE IN OUR COMPUTER TO FIX THE ISSUE
# GO TO THIS LINK: https://www.microsoft.com/en-us/download/confirmation.aspx?id=13255
# YOU HAVE TWO FILES TO DOWNLOAD, CHOOSE (AccessDatabaseEngine_X64.exe) AND HIT DOWNLOAD
# WHEN IT IS DONE, INSTALL THE FILE ON YOUR COMPUTER, MOST LIKELY IT WILL GIVE YOU AN ERROR OF (YOU HAVE 64-BIT ,
# CAN'T DOWNLOAD 32-BIT)
# TO OVER THAT OPEN COMMAND PROMPT FROM START MENU
# ACCESS THE FILE IN DOWNLOAD FOLDER BY CHANGE DIRECTORY, JUST PASTE THAT: C:\Users\malatweh\Downloads
# WHEN YOU ARE IN RIGHT FOLDER PASTE THAT: AccessDatabaseEngine_X64.exe /passive
# (/passive): USED TO FIX MISMATCH VERSION ISSUE AND FORCE WINDOWS TO INSTALL THE FILE
# MOST LIKELY YOU NEED TO RESTART YOUR COMPUTER

# TRY THE FOLLOWING IN PYTHON EDITOR
# import pyodbc
# print([x for x in pyodbc.drivers() if x.startswith('Microsoft')])

# IF IT PRINT:

# ['Microsoft Access Driver (*.mdb, *.accdb)',
#  'Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)',
#  'Microsoft Access dBASE Driver (*.dbf, *.ndx, *.mdx)',
#  'Microsoft Access Text Driver (*.txt, *.csv)']

# ---> YOU ARE READY TO GO

# RESOURCES
# https://stackoverflow.com/questions/45928987/is-it-possible-for-64-bit-pyodbc-to-talk-to-32-bit-ms-access-database
# https://www.microimages.com/downloads/MS_AccessDB.htm
# https://datasavvy.me/2017/07/20/installing-the-microsoft-ace-oledb-12-0-provider-for-both-64-bit-and-32-bit-processing/
# https://social.msdn.microsoft.com/Forums/office/en-US/57aee87f-a2e0-4804-a452-7c69f1d32957/how-i-can-install-access-database-engine-without-uninstalling-microsoft-office?forum=exceldev


# USE (pyodbc) MODULE TO CONNECT PYTHON WITH MICROSOFT ACCESS
import pyodbc
# TO CONNECT PYTHON TO MICROSOFT ACCESS FILE JUST PUT THE PATH OF THE FILE AFTER (DBQ=)
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=')
cursor = conn.cursor()
# TO SHOW EVERYTHING INSIDE THE TABLE DATABASE USE (*) TO SELECT EVERYTHING FROM THE TABLE(dbo_tblForgeSpecAll)
cursor.execute('select * from dbo_tblForgeSpecAll')
for row in cursor.fetchall():
    # print(row)
    pass
print()
# TO SHOW SPECIFIC COLUMN INSIDE THE TABLE DATABASE USE CULOMN NAME (vcrIdentifier) TO SELECT COLUMN DATA FROM
# THE TABLE(dbo_tblForgeSpecAll)
# (fetchall()): USED TO SHOW EVERYTHING RELATE IN DATA WE SEARCHED FOR
# (fetchone()): USED TO SHOW JUST THE NUMBER OF THE DATA WE SEARCHED FOR>>> WE USE THIS ONE BECAUSE WE NEED
# JUST THE WITHOUT ANY EXTRA INFORMATION
cursor.execute('select vcrIdentifier from dbo_tblForgeSpecAll')
# cursor.execute('select * from dbo_tblForgeSpecAll WHERE vcrIdentifier')
for row in cursor.fetchall():
    print(row)
print()
# TO SHOW ALL DATA FOR SPECIFIC FORGING NUMBER
# (*): TO SELECT EVERYTHING
# (from dbo_tblForgeSpecAll): TO ACCESS TABLE OF DATABASE
# (WHERE vcrIdentifier='F6444X') : FROM COLUMN 'vcrIdentifier' ACCESS 'F6444X'
# USE DOUBLE QUOTE("") FOR WHOLE THING, AND SINGLE QUOTE('') FOR FORGING NUMBER
# (fetchall()): USED TO SHOW EVERYTHING RELATE IN DATA WE SEARCHED FOR
# (fetchone()): USED TO SHOW JUST THE NUMBER OF THE DATA WE SEARCHED FOR>>> WE USE THIS ONE BECAUSE WE NEED
# JUST THE WITHOUT ANY EXTRA INFORMATION
cursor.execute("select * from dbo_tblForgeSpecAll WHERE vcrIdentifier='F6444X'")
for row in cursor.fetchone():
    print(row)
print()
# TO SHOW SPECIFIC DATA FOR SPECIFIC FORGING NUMBER
# (decForgeRefLength): TO CHOOSE COLUMN THAT CONTAIN DATA YOU WANT,
# EX:(decForgeRefLength) STORE '(F)Forge Ref Length' FOR EACH FORGING>>
# >>CHANGE COLUMN TO WHATEVER DATA YOU WANT LIKE '(B)Forge O.D.','(J)Boss Insd Spacing'...ETC
# (from dbo_tblForgeSpecAll): TO ACCESS TABLE OF DATABASE
# (WHERE vcrIdentifier='F6444X') : FROM COLUMN 'vcrIdentifier' ACCESS 'F6444X'
# (fetchall()): USED TO SHOW EVERYTHING RELATE IN DATA WE SEARCHED FOR
# (fetchone()): USED TO SHOW JUST THE NUMBER OF THE DATA WE SEARCHED FOR>>> WE USE THIS ONE BECAUSE WE
# NEED JUST THE WITHOUT ANY EXTRA INFORMATION
print(cursor.execute("select decForgeRefLength from dbo_tblForgeSpecAll WHERE vcrIdentifier='F6444X'"))
for row in cursor.fetchone():
    # WE CAN DO THAT TO STORE THE FOUND DATA IN VARIABLE THAT WILL USE TO BUILD THE APP
    ForgeRefLength = row
    print(ForgeRefLength)
