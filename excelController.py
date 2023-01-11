import os
import openpyxl


def getPath():

    # Give the location of the file
    osPath = os.getcwd() #get path of this directory
    file = '\delivery.xlsx' #filename
    return osPath+file


def openExcel():
    
    path = getPath()

    # To open the workbook
    # workbook object is created
    workbook = openpyxl.load_workbook(path)
    print('file is openend')
    return workbook
    
def readDataFromExcel(workbook, data):
    
    # Get sheet names
    sheet = workbook.active
    
    # Read a Cell
    # Get cell value
    cell_obj = sheet.cell(row=1, column=1)
    
    # Print value of cell
    print(cell_obj.value)
    
def deleteArticleExpiration():
    print('delete Article expiration')

def writeDataToExcel(workbook, articleObjectList):
    
    # Get sheet names
    sheet = workbook.active
    
    for articleObject in articleObjectList:
        max_row = sheet.max_row # find last row of worksheet
        # Write to a Cell
        # write to a particular cell
        sheet.cell(row= max_row+1, column=1).value = articleObject.bestelbonnr
        sheet.cell(row= max_row+1, column=2).value = articleObject.articleId, 
        sheet.cell(row= max_row+1, column=3).value = articleObject.number
        sheet.cell(row= max_row+1, column=4).value = articleObject.deliveryDate

    path = getPath()
    # Save the file
    workbook.save(path)