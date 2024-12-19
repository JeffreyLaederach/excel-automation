import xlsxwriter

# Create a new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook("Spreadsheets/Excel Python Test.xlsx")
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column("A:A", 30)

# Insert an image:
worksheet.write("A1", "This is Cell A1")
# worksheet.insert_image("A2", "example.png")

workbook.close()