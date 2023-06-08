import openpyxl
from datetime import datetime

# Creating a new file to generate the table into

wb = openpyxl.Workbook()
# Creating filename with full length date string to avoid collisions (checking not necessary)
filename = datetime.now().strftime("%d%m%Y %H-%M-%S")
# Save empty workbook initially to establish name is not used and getting a file handle
wb.save(filename)

