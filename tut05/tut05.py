from datetime import datetime
start_time = datetime.now()

###importing input excel file
import openpyxl
wb = openpyxl.load_workbook(r'octant_input.xlsx')
sheet = wb.active

###calculating no. of rows
count_row=sheet.max_row
t_c=count_row-1


### List for storing octant signs
oct_sign_lst = [1, -1, 2, -2, 3, -3, 4, -4]
