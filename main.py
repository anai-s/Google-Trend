from pytrends.request import TrendReq
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle


pytrend = TrendReq()
geo = input('Choose your geolocalization (ex:FR): ')
search_terms = input('Choose your terms (use space as separator ex:bricolage decoration jardin): ')
kw_list = search_terms.split()
timeframe = input('Choose your timeframe (ex:today 1-m): ')
pytrend.build_payload(kw_list = kw_list, geo = geo, timeframe = timeframe)
related_queries = pytrend.related_queries()
search_terms = [related_queries.get(elem).get('rising') for elem in kw_list]

for i in range(len(search_terms)):
    search_terms[i].insert(loc=0, column="theme", value=kw_list[i])
    
result = pd.concat(search_terms).reset_index()

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# here enter the id of your google sheet
SAMPLE_SPREADSHEET_ID_INPUT = 'YOUR GOOGLE SHEET ID'
SAMPLE_RANGE_NAME = 'A1:AA1000'
CLIENT_SECRET_SERVICE = 'YOUR CLIENT ID JSON'
 
def Create_Service(client_secret_file, api_service_name, api_version, scopes):

    cred = None
 
    if os.path.exists('token_write.pickle'):
        with open('token_write.pickle', 'rb') as token:
            cred = pickle.load(token)
 
    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, scopes)
            cred = flow.run_local_server()
 
        with open('token_write.pickle', 'wb') as token:
            pickle.dump(cred, token)
 
    try:
        service = build(api_service_name, api_version, credentials=cred)
        #print(api_service_name, 'service created successfully')
        return service
    except Exception as e:
        print(e)
        return None

service = Create_Service(CLIENT_SECRET_SERVICE, 'sheets', 'v4', SCOPES)
     
def Export_Data_To_Sheets():
    service.spreadsheets().values().batchClear(spreadsheetId=SAMPLE_SPREADSHEET_ID_INPUT,
    body= { "ranges":[SAMPLE_RANGE_NAME]}).execute()
    service.spreadsheets().values().update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID_INPUT,
        valueInputOption='RAW',
        range=SAMPLE_RANGE_NAME,
        body=dict(
            majorDimension='ROWS',
            values=result.T.reset_index().T.values.tolist())
    ).execute()
    print('Sheet successfully Updated')
 
Export_Data_To_Sheets()