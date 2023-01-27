import os
import json
from datetime import date


def getPath():

    # Give the location of the file
    osPath = os.getcwd()  # get path of this directory
    file = '\data.json'  # filename
    return osPath+file


def lastDateRun():

    path = getPath()

    # Opening JSON file
    with open(path, 'r') as openfile:

        # Reading from json file
        jsonObject = json.load(openfile)
        lastRun = jsonObject['data']['lastRun']
        return lastRun


def updateLastDateRun():

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
