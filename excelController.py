import os
import openpyxl
from datetime import date, datetime


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


# This function delete all the rows in excel with the same ordernummer and artikelcode that was found in de emails
def deleteArticleDubbel(workbook, articleObjectList):

    # Get sheet names
    sheet = workbook.active
    index = 1
    del_rows = []

    for row in sheet.iter_rows(min_row=2):
        index += 1
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


def deleteArticlesRowExpireddate(workbook):

    today = date.today()
    print(str(today))

    dateTimeNow = datetime.now()
    print(str(dateTimeNow))

    # Get sheet names
    sheet = workbook.active

    index = 1
    del_rows = []

    # there is an issue with the first row. because is a tekst and not a Date
    # So this loop has to start from the second row.
    for row in sheet.iter_rows(min_row=2):
        index += 1
        deliverDateStr = row[3].value
        deliverDate = datetime.strptime(deliverDateStr, "%d/%m/%y")

        if dateTimeNow > deliverDate:
            # If deliverDate is expired with current day
            del_rows.append(index)

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
