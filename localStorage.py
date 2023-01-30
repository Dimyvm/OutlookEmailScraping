import os
import json
from datetime import date
from errorHandler import *


def getPath():

    # Give the location of the file
    osPath = os.getcwd()  # get path of this directory
    file = '\data.json'  # filename
    return osPath+file


def lastDateRun():

    try:
        path = getPath()

        # Opening JSON file
        with open(path, 'r') as openfile:

            # Reading from json file
            jsonObject = json.load(openfile)
            lastRun = jsonObject['data']['lastRun']
            return lastRun
    except Exception as e:

        title = 'De code is vastgelopen in de lastDateRun functie'
        error = e
        SendErrorMail(title, error)
        input()


def updateLastDateRun():
    try:
        path = getPath()
        today = date.today()

        data = {
            "data": {
                "lastRun": str(today)
            }
        }

        json_string = json.dumps(data)
        with open(path, 'w', encoding='utf-8') as outfile:
            outfile.write(json_string)

    except Exception as e:

        title = 'De code is vastgelopen in de updateLastDateRun functie'
        error = e
        SendErrorMail(title, error)
        input()
