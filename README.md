import pandas as pd
import os
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Read Excel files
df1 = pd.read_excel(r"C:\Users\Lenovo\Downloads\TestTest.xlsx", sheet_name=0)
df2 = pd.read_excel(r"C:\Users\Lenovo\Downloads\TestTest.xlsx", sheet_name=1)

# load the workbook
wb = load_workbook(filename=r"C:\Users\Lenovo\Downloads\TestTest.xlsx", data_only=True)

# select a sheet
sheet = wb['Tabelle1']
sheet2 = wb['Tabelle2']

# initialize an empty list
values = []
valuesCellNumber = []
valuesWords = []
lastCellValue = []
undercardValue = []

# Initialize an empty list for the dictionaries
linked_data = []

# Iterate through the 'A' column
for i in range(sheet.max_row):
    cell_value = sheet['A' + str(i+1)].value
    # check if the cell is not empty
    if cell_value is not None:
        # append the value to the list
        values.append(sheet['A' + str(i+1)].value)
        valuesCellNumber.append(i+1)

for val in valuesCellNumber:
    last_column = sheet.max_column
    last_cell_value = sheet.cell(row=val, column=last_column).value
    lastCellValue.append(last_cell_value)

# Iterate through the 'C' column
for i in range(sheet.max_row):
    cell_valueWord = sheet['C' + str(i+1)].value
    # check if the cell is not empty
    # append the value to the list
    valuesWords.append(sheet['C' + str(i+1)].value)

# Get the last column letter
last_column_letter = get_column_letter(sheet.max_column)

# Iterate through the rows in the last column
for i in range(1, sheet.max_row + 1):
    cell_value = sheet[last_column_letter + str(i)].value
    # check if the cell is not empty
    # append the value to the list
    undercardValue.append(cell_value)

print(undercardValue)
print(values)
print(lastCellValue)

# Initialize the current id and number
current_id_index = 0
current_number_index = 0
words = []

# Iterate over the words
for word in valuesWords:
    # If word is None, move to the next id and number
    if word is None:
        linked_data.append({'id': values[current_id_index], 'word': words, 'number': lastCellValue[current_number_index]})
        words = []
        current_id_index += 1
        current_number_index += 1
    else:
        words.append(word)

# Add the last group if words is not empty
if words:
    linked_data.append({'id': values[current_id_index], 'word': words, 'number': lastCellValue[current_number_index]})

# Now linked_data is a list of dictionaries, where each dictionary represents a row in the Excel table
print(linked_data)

# Iterate over the words again
for word in valuesWords:
    # If word is None, move to the next id and number
    if word is None:
        linked_data.append({'id': values[current_id_index], 'word': words, 'number': lastCellValue[current_number_index]})
        words = []
        current_id_index += 1
        current_number_index += 1
    else:
        words.append(word)

# Add the last group if words is not empty
if words:
    linked_data.append({'id': values[current_id_index], 'word': words, 'number': lastCellValue[current_number_index]})

# Now linked_data is a list of dictionaries, where each dictionary represents a row in the Excel table
print(linked_data)

# Iterate through the rows under the last cell value
for i in range(len(lastCellValue) + 1, sheet.max_row + 1):
    cell_value = sheet[last_column_letter + str(i)].value
    # If the cell value is None, break the loop
    if cell_value is None:
        break
    # Otherwise, append the cell value to the list and print it
    undercardValue.append(cell_value)
    print(cell_value)
