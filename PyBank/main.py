import os
import csv
import xlsxwriter

with open('budget.csv', newline='') as csvfile:
    csvreader = csv.reader(csvfile, delimiter=',')
    next(csvreader, None)

    total = 0
    months = []

    print(f"Financial Analysis")
    print(f"--------------------------------------------")

    for row in csvreader:
        months.append(row[0])
        total = total + int(row[1])

    total_months = len(months)

    print(f"Total Months: {total_months}")
    print(f"Total: ${total:,}")

with open('budget.csv', newline='') as csvfile2:

    csvreader2 = csv.reader(csvfile2, delimiter=',')
    next(csvreader2, None)

    csvlist = list(csvreader2)
    first = int(csvlist[0][1])
    last = int(csvlist[-1][1])
    change = last - first
    average_change = change / (len(months)-1)

    print(f"Average Change: ${average_change:.2f}")

with open('budget.csv', newline='') as csvfile3:

    csvreader3 = csv.reader(csvfile3, delimiter=',')
    next(csvreader3, None)

    x = int(row[1])
    y = 0
    changelist = []
    dateslist = []

    for row in csvreader3:
        if (x - y) == x:
            x = int(row[1])
            dateslist.append(row[0])
            changelist.append(0)
            y = x
        else:
            x = int(row[1])
            dateslist.append(row[0])
            changelist.append(x - y)
            y = x

    changedict = dict(zip(dateslist,changelist))
 
    maximum = max(changedict, key=changedict.get)
    minimum = min(changedict, key=changedict.get)

    print(f"Greatest Increase in Profits: {maximum} ($ {changedict[maximum]:,})")
    print(f"Greatest Decrease in Profits: {minimum} (${changedict[minimum]:,})")

workbook = xlsxwriter.Workbook("analysis.xlsx")
worksheet = workbook.add_worksheet()

bold = workbook.add_format({"bold":1})
num_format = workbook.add_format({"num_format": "#,##0.00"})
worksheet.set_column("A:A", 30)
worksheet.set_column("B:B", 15)

worksheet.write("A1", "Financial Analysis", bold)
worksheet.write("A2", "Total Months:")
worksheet.write("A3", "Total:")
worksheet.write("A4", "Average Change:")
worksheet.write("A5", "Greatest Increase in Profits:")
worksheet.write("A6", "Greatest Decrease in Profits:")

row = 0
col = 1

worksheet.write_number(row + 1, col, total_months)
worksheet.write_number(row + 2, col, total, num_format)
worksheet.write_number(row + 3, col, average_change, num_format)
worksheet.write_number(row + 4, col, changedict[maximum], num_format)
worksheet.write_string(row + 4, col + 1, maximum, num_format)
worksheet.write_number(row + 5, col, changedict[minimum], num_format)
worksheet.write_string(row + 5, col + 1, minimum, num_format)

workbook.close()