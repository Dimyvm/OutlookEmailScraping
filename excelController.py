import os
import openpyxl


def getPath():

    # Give the location of the file
    osPath = os.getcwd()  # get path of this directory
    file = '\delivery.xlsx'  # filename
    return osPath+file


def openExcel():

    path = getPath()

    # To open the workbook
    # workbook object is created
    workbook = openpyxl.load_workbook(path)
    return workbook


def readDataFromExcel(workbook, data):

    # Get sheet names
    sheet = workbook.active

    # Read a Cell
    # Get cell value
    cell_obj = sheet.cell(row=1, column=1)

    # Print value of cell
    print(cell_obj.value)


# This function has yet to be tested !
def deleteArticleExpiration(workbook, articleObjectList):

    # Get sheet names
    sheet = workbook.active
    index = 1
    del_rows = []

    for row in sheet.iter_rows():

        orderNumber = row[0].value
        articleId = row[1].value
        number = row[2].value

        for articleObject in articleObjectList:
            if orderNumber == articleObject.orderNumber and articleId == articleObject.articleId and number == articleObject.number:
                # row matches object
                del_rows.append(index)
                break

    for r in reversed(del_rows):
        sheet.delete_rows(r)

    path = getPath()
    # Save the file
    workbook.save(path)


def writeDataToExcel(workbook, articleObjectList):

    # Get sheet names
    sheet = workbook.active

    for articleObject in articleObjectList:
        max_row = sheet.max_row  # find last row of worksheet
        # Write to a Cell
        # write to a particular cell
        sheet.cell(row=max_row+1, column=1).value = articleObject.orderNumber
        sheet.cell(row=max_row+1, column=2).value = articleObject.articleId
        sheet.cell(row=max_row+1, column=3).value = articleObject.number
        sheet.cell(row=max_row+1, column=4).value = articleObject.deliveryDate
        max_row = max_row+1

    path = getPath()
    # Save the file
    workbook.save(path)
