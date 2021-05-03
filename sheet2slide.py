from __future__ import print_function

from apiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools

SCOPES = (
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/presentations',
)
store = file.Storage('storage.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('client_secrets.json', SCOPES)
    creds = tools.run_flow(flow, store)
HTTP = creds.authorize(Http())
SHEETS = discovery.build('sheets', 'v4', http=HTTP)
SLIDES = discovery.build('slides', 'v1', http=HTTP)

print('** Fetching Sheet Data')
sheetID = 'INSERT HERE YOUR SHEET ID'
orders = SHEETS.spreadsheets().values().get(range='Página1',
        spreadsheetId=sheetID).execute().get('values')

print('** Fetching Sheet Chart Info')
sheet = SHEETS.spreadsheets().get(spreadsheetId=sheetID,
        ranges=['Insert Here Your Page']).execute().get('sheets')[0]
chartID = sheet['charts'][0]['chartId']

print('** Creating a New Slide Deck')
DATA = {'title': 'API Generated Presentation'}
rsp = SLIDES.presentations().create(body=DATA).execute()
deckID = rsp['presentationId']
titleSlide = rsp['slides'][0]
titleID = titleSlide['pageElements'][0]['objectId']
subtitleID = titleSlide['pageElements'][1]['objectId']

print('** Creating Slides and Inserting Title/Subtitle')
reqs = [
    {'createSlide': {'slideLayoutReference': {'predefinedLayout': 'TITLE_ONLY'}}}, #Create slide with title to insert, slide 0
    {'createSlide': {'slideLayoutReference': {'predefinedLayout': 'BLANK'}}}, #Create a blank slide
    {'insertText': {'objectId': titleID, 'text': 'Insert here your TITLE'}},
    {'insertText': {'objectId': subtitleID, 'text': 'Insert here your SUBTITLE'}},
]
rsp = SLIDES.presentations().batchUpdate(body={'requests': reqs},
        presentationId=deckID).execute().get('replies')
tableSlideID = rsp[0]['createSlide']['objectId']
chartSlideID = rsp[1]['createSlide']['objectId']

print('** Fetching Table Title')
rsp = SLIDES.presentations().pages().get(presentationId=deckID,
        pageObjectId=tableSlideID).execute().get('pageElements')
textboxID = rsp[0]['objectId']

print('** Creating the Table and inserting Title')
reqs = [
    {'createTable': {
        'elementProperties': {'pageObjectId': tableSlideID},
        'rows': len(orders),
        'columns': len(orders[1])}
    },
    {'insertText': {'objectId': textboxID, 'text': 'Atividades'}},
]
rsp = SLIDES.presentations().batchUpdate(body={'requests': reqs},
    presentationId=deckID).execute().get('replies')
tableID = rsp[0]['createTable']['objectId']

print('** Completing Cells from the Table and Linking Sheet Chart with Slide')
reqs = [
    {'insertText': {
        'objectId': tableID,
        'cellLocation': {'rowIndex': i, 'columnIndex': j},
        'text': str(data),    
    }} for i, order in enumerate(orders) for j, data in enumerate(order)]

reqs.append({'createSheetsChart': {
    'spreadsheetId': sheetID,
    'chartId': chartID,
    'linkingMode': 'LINKED',
    'elementProperties': {
        'pageObjectId': chartSlideID,
        'size': {
            'height': {'magnitude': 7075, 'unit': 'EMU'}, #Dimensions of Chart Size
            'width': {'magnitude': 11450, 'unit': 'EMU'}
        },
        'transform': {
            'scaleX': 696.6157,
            'scaleY': 601.3921,
            'translateX': 583875.04,
            'translateY': 444327.135,
            'unit': 'EMU',
        },
    },
}})
SLIDES.presentations().batchUpdate(body={'requests': reqs},
        presentationId=deckID).execute()
print('Done')
