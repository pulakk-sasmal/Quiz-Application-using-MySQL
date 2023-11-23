import openpyxl
import MySQLdb

loc = "Resources/Questions.xlsx"

wb = openpyxl.load_workbook(loc)
sheet = wb.active
n = sheet.max_row

conn = MySQLdb.connect(host='localhost', database = 'world', user='pulak', password='123123')
cursor = conn.cursor()

cursor.execute("DROP TABLE IF EXISTS Questions")
q = "CREATE TABLE questions(QID INT, qstn TEXT, opA TEXT, opB TEXT, opC TEXT, opD TEXT, ans INT)"
cursor.execute(q)

# Code to add rows
for i in range(2, n + 1):  # Start from the second row (header row is skipped)
    q2 = "INSERT INTO questions(QID, qstn, opA, opB, opC, opD, ans) VALUES (%s, %s, %s, %s, %s, %s, %s)"
    lst = [cell.value for cell in sheet[i]]
    arg = (lst[0], lst[1], lst[2], lst[3], lst[4], lst[5], lst[6])
    cursor.execute(q2, arg)
    conn.commit()

cursor.close()
conn.close()
