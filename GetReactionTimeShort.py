import os, glob, xlsxwriter

workbook = xlsxwriter.Workbook('ReactionTime.xlsx') #Creates an Excel file in your current folder

worksheet = workbook.add_worksheet("CPT")

description = ["ID", "Age", "Sex", "Trial 1", "Trial 2", "Trial 3"]
row = 0
column = 0 

for item in description:
    worksheet.write(row, column, item)
    column += 1

row = 1
column = 0

path = "C:/Users/felix/Desktop/Hovedoppgave/Eksempelfiler" #Should be the path where your files are

for filename in glob.glob(os.path.join(path, "*.txt")):
    with open(filename, "r+") as fo:
        for i, line in enumerate(fo):
            if i == 7:
                worksheet.write(row, column, (line[11:18]).strip())
                column += 1
            if i == 8:
                worksheet.write(row, column, (line[5:10]).strip())
                column += 1
            if i == 9:
                worksheet.write(row, column, (line[5:10]).strip())
                column += 1
            if i == 18:
                worksheet.write(row, column, (line[22:31]).strip())
                column += 1
            if i == 31:
                worksheet.write(row, column, (line[22:30]).strip())
                column += 1
            if i == 44:
                worksheet.write(row, column, (line[22:30]).strip())
                column = 0 
                row += 1
        fo.close()

workbook.close()