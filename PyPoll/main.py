import os
import csv
import xlsxwriter

with open("election.csv", newline="") as csvfile:
    csvreader = csv.reader(csvfile, delimiter=",")
    next(csvreader, None)

    votes = []

    for row in csvreader:
        votes.append(row[0])

    total_votes = len(votes)

print("Election Results")  
print("------------------------------------")  
print(f"Total Votes: {total_votes:,}")
print("------------------------------------")  

with open("election.csv", newline="") as csvfile2:
    csvreader2 = csv.reader(csvfile2, delimiter=",")
    next(csvreader2, None)

    candidates = []

    for row in csvreader2:
        if row[2] not in candidates:
            candidates.append(row[2])

print(f"The Candidates are: {candidates}")
print("------------------------------------")  

with open("election.csv", newline="") as csvfile3:
    csvreader3 = csv.reader(csvfile3, delimiter=",")
    next(csvreader3, None)

    a = candidates[0]
    b = candidates[1]
    c = candidates[2]
    d = candidates[3]
    vcountera = 0
    vcounterb = 0
    vcounterc = 0
    vcounterd = 0

    for row in csvreader3:
        if row[2] == a:
            vcountera = vcountera + 1
        elif row[2] == b:
            vcounterb = vcounterb + 1
        elif row[2] == c:
            vcounterc = vcounterc + 1
        elif row[2] == d:
            vcounterd = vcounterd + 1

votesdict = {a: vcountera, b: vcounterb, c: vcounterc, d: vcounterd}
maximum = max(votesdict, key=votesdict.get)

print(f"{a} : {vcountera / total_votes:.2%} ({vcountera:,})")
print(f"{b} : {vcounterb / total_votes:.2%} ({vcounterb:,})")
print(f"{c} : {vcounterc / total_votes:.2%} ({vcounterc:,})")
print(f"{d} : {vcounterd / total_votes:.2%} ({vcounterd:,})")
print("------------------------------------") 
print(f"The winner is: {maximum} with {votesdict[maximum]:,} votes")

workbook = xlsxwriter.Workbook("results.xlsx")
worksheet = workbook.add_worksheet()

bold = workbook.add_format({"bold":1})
num_format = workbook.add_format({"num_format": "#,##0"})
worksheet.set_column("A:A", 30)
worksheet.set_column("B:B", 15)
worksheet.set_column("C:C", 15)

worksheet.write("A1", "Election Results", bold)
worksheet.write("A2", "Total Votes:")
worksheet.write("A3", "Candidate 1:")
worksheet.write("A4", "Candidate 2:")
worksheet.write("A5", "Candidate 3:")
worksheet.write("A6", "Candidate 4:")
worksheet.write("A7", "The winner is:", bold)

row = 0
col = 1

worksheet.write_number(row + 1, col, total_votes, num_format)
worksheet.write_string(row + 2, col, a)
worksheet.write_number(row + 2, col + 1, vcountera, num_format)
worksheet.write_string(row + 3, col, b)
worksheet.write_number(row + 3, col + 1, vcounterb, num_format)
worksheet.write_string(row + 4, col, c)
worksheet.write_number(row + 4, col + 1, vcounterc, num_format)
worksheet.write_string(row + 5, col, d)
worksheet.write_number(row + 5, col + 1, vcounterd, num_format)
worksheet.write_string(row + 6, col, maximum)
worksheet.write_number(row + 6, col + 1, votesdict[maximum], num_format)

workbook.close()