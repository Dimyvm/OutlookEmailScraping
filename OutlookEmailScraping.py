import win32com.client
import win32ui
from datetime import date, timedelta, datetime
from localStorage import *
from articleClass import Article
from excelController import *
from errorHandler import *


def main():
    try:
        # def read whene it was the last time this script was executed
        lastRun = lastDateRun()
        print(
            f'This script was run successfully for the last time on {lastRun}')

        # read outlook emails
        messages = readOutlookMails()

        # filter messages
        filteredMessages = filtermails(messages, lastRun)
        messagesCount = len(filteredMessages)

        # get data from filtered mails
        articleObjectList = readDatafromMail(filteredMessages)
        articleObjectListLength = len(articleObjectList)
        print(f'The length of the articlelist is: {articleObjectListLength}')
        print(f"There are {messagesCount} mails found with the right subject")

        # def open excel
        workbook = openExcel()

        # delete each dubble record where the deliverydate == to the date of Today
        deleteArticleDubbel(workbook, articleObjectList)
        # Write each article and deliverydate in a excel record
        writeDataToExcel(workbook, articleObjectList)
        # delete each record where the deliverydate is older to the date of Today
        deleteArticlesRowExpireddate(workbook)
        checkOnDubbel(workbook)

    except Exception as e:

        title = 'De code is vastgelopen in de Main functie'
        error = e
        SendErrorMail(title, error)
        input()

    else:
        # if script is executed successfully then show windows messageBox
        win32ui.MessageBox("""Successfully executed!""",
                           "Siemens email scraping")
        # write the date whene this script was executed
        updateLastDateRun()


def readOutlookMails():
    try:
        # def read outlook emails
        outlook = win32com.client.dynamic.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        # "6" refers to the index of a folder - in this case,
        # "5" = Verzonden items
        messages = inbox.Items
        return messages

    except Exception as e:
        title = 'De code is vastgelopen in de readOutlookMails functie'
        error = e
        SendErrorMail(title, error)


def filtermails(messages, lastRun):
    try:
        # def filter emails
        messages.Sort("[ReceivedTime]", True)  # sort messages on date
        lastRunPlusOne = datetime.strptime(
            lastRun, "%Y-%m-%d") + timedelta(days=1)
        filteredMessages = messages.Restrict(
            "[ReceivedTime] >= '" + lastRunPlusOne.strftime('%d/%m/%Y %H:%M %p')+"'")
        print(
            f"There are {filteredMessages.count} messages collected from {lastRun} to today")
        return filteredMessages

    except Exception as e:
        title = 'De code is vastgelopen in de filtermails functie'
        error = e
        SendErrorMail(title, error)


def readDatafromMail(filteredMessages):
    try:
        # read the data from each filtered email
        aantal = 0
        articleObjectList = []
        for message in filteredMessages:
            if message.subject.startswith('SIEMENS - Update'):
                aantal += 1
                body_title = message.subject
                body_content = message.body
                sendDate = message.SentOn.strftime("%d-%m-%y")

                print(f"Send date : {sendDate}")
                bestelbonnr = body_title.split("/")[0][-7:]
                print(f'Bestelbonnr:{bestelbonnr}')

                lines = body_content.splitlines()

                print(bestelbonnr)
                articleId = ""
                deliveryDate = ""
                number = ""

                for line in lines:
                    if "Klantartikel" not in line:
                        if line.startswith("ArtikelID") or line.startswith("Artikelnummer"):
                            if line.split()[1] != 'klant':
                                articleId = line.split()[1]
                        if line.startswith("Bevestigde leverdatum"):
                            deliveryDate = line.split()[2]
                        if line.startswith("Bevestigd aantal"):
                            number = line.split()[2]
                            print(articleId)
                            print(deliveryDate)
                            print(number)

                            articleObjectList.append(
                                Article(bestelbonnr, articleId, number, deliveryDate))
                            print("article is added")
                            print("----------")
        return articleObjectList

    except Exception as e:
        title = 'De code is vastgelopen in de readDatafromMail functie'
        error = e
        SendErrorMail(title, error)


main()
