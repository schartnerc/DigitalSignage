import os
import requests
import json
import time
import csv
import configparser
from collections import defaultdict

from oauth2client.client import OAuth2WebServerFlow
from oauth2client.tools import run_flow
from oauth2client.file import Storage


# Gets access token from auth server.
def get_access_token():
    credentials = Storage(config['SecretFile']).get()
    if credentials is None or credentials.id_token['exp'] <= int(time.time()):
        flow = OAuth2WebServerFlow(client_id=config['ClientID'],
                                   client_secret=config['ClientSecret'],
                                   scope=config['Scope'],
                                   redirect_uri=['RedirectURI'],
                                   login_hint=config['User'])
        storage = Storage(filename=config['SecretFile'])
        credentials = run_flow(flow, storage)
    return credentials.access_token


# Executes the API request
def api_request(category):
    url = config['APIUri'] + category
    querystring = {"companyId": config['CompanyID'],
                   "count": config['Limit'],
                   "access_token": get_access_token()}

    response = requests.request("GET", url, params=querystring)
    if response.status_code is 200:
        return json.loads(response.text)
    else:
        raise ValueError('Authentication failed')


def _check_for_key(_data, key):
    if key in _data:
        return _data[key]
    else:
        return 'unlimited'


# Extracts data from the api response and
def get_excel_data():
    data = api_request('schedules')
    sign_data = defaultdict(dict)
    presentations = set()
    signs = set()

    for sign in data['items']:
        for presentation in sign['content']:
            sign_data[presentation['name']].update({
                sign['name']:
                    {
                        'startDate': _check_for_key(presentation, 'startDate'),
                        'endDate': _check_for_key(presentation, 'endDate')
                    }
                })
            presentations.add(presentation['name'])
            signs.add(sign['name'])

    presentations = list(presentations)
    presentations.sort()

    signs = list(signs)
    signs.sort()

    with open('output.csv', 'w') as output:
        writer = csv.writer(output, delimiter=';', quotechar='"')
        csv_header = ['']
        for sign in signs:
            csv_header.append(str(sign + '_start_date'))
            csv_header.append(str(sign + '_end_date'))
        writer.writerow(csv_header)
        for presentation in presentations:
            csv_row = [str(presentation)]
            for sign in signs:
                if sign in sign_data[presentation]:
                    item = sign_data[presentation][sign]
                    csv_row.append(str(item['startDate']))
                    csv_row.append(str(item['endDate']))
                else:
                    csv_row.append('')
                    csv_row.append('')
            writer.writerow(csv_row)


if __name__ == '__main__':
    config = configparser.ConfigParser()
    config.read(os.path.join(os.path.dirname(__file__), 'parameters.ini'))
    config = config['RiseVision App']
    try:
        get_excel_data()
    except ValueError:
        print('User authentication failed')





