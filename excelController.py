import os
import openpyxl


def openExcel():
    # Give the location of the file
    osPath = os.getcwd()
    print("Current working directory:", str(osPath))
    file = '\delivery.xlsx'

    # To open the workbook
    # workbook object is created
    workbook = openpyxl.load_workbook(osPath+file)

    # Get sheet names
    sheet = workbook.active

    # Get current active sheet
    # from the active attribute
    # sheet.title

    # Read a Cell
    # Get cell value
    cell_obj = sheet.cell(row=1, column=1)

    # Print value of cell
    print(cell_obj.value)

    # Write to a Cell
    # write to a particular cell
    sheet.cell(row=2, column=2).value = "Welcome"

    # Save the file
    workbook.save(osPath+file)
