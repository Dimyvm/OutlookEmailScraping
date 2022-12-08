import json
from datetime import date


def lastDateRun():

    # Opening JSON file
    with open('data.json', 'r') as openfile:

        # Reading from json file
        jsonObject = json.load(openfile)
        lastRun = jsonObject['data']['lastRun']
        return lastRun


def updateLastDateRun():
    today = date.today()
    data = {
        "data": {
            "lastRun": str(today)
        }
    }

    json_string = json.dumps(data)
    with open('data.json', 'w', encoding='utf-8') as outfile:
        outfile.write(json_string)
