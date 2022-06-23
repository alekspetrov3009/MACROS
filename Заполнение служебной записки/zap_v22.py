from openpyxl import Workbook
import datetime

excel_file = Workbook()
excel_sheet = excel_file.create_sheet(title='Holidays 2019', index=0)

# creating header row
excel_sheet['A1'] = 'Holiday Name'
excel_sheet['B1'] = 'Holiday Description'
excel_sheet['C1'] = 'Holiday Date'

# adding data
excel_sheet['A2'] = 'Diwali'
excel_sheet['B2'] = 'Biggest Indian Festival'
excel_sheet['C2'] = datetime.date(year=2019, month=10, day=27).strftime("%m/%d/%y")

excel_sheet['A3'] = 'Christmas'
excel_sheet['B3'] = 'Birth of Jesus Christ'
excel_sheet['C3'] = datetime.date(year=2019, month=12, day=25).strftime("%m/%d/%y")

# save the file
excel_file.save(filename="Holidays.xlsx")