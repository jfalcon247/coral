from __future__ import print_function
import httplib2
import os
import numpy as np
import pandas as pd
from googleapiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/presentations',
          'https://www.googleapis.com/auth/drive.metadata'
         ]
         
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'coralgardeners'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    credential_path = '/Users/jfalcon/Desktop/coral/client_secret.json'
    print(credential_path)

    store = Storage('storage.json')
    credentials = store.get()

    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def makeCertificate(orderNum, name, gpsLat, gpsLong, instagram, slideFlagAddress, pdfFlagAddress):

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    # initialize API services
    SHEETS = discovery.build('sheets', 'v4', http=http)
    SLIDES = discovery.build('slides', 'v1', http=http)
    DRIVE = discovery.build('drive', 'v3', http=http)

    template_id = '1feAh6Ylguu3gLuvs4z5aQslTiICovJHRKkM4Fsutri8'
    sheet_id = '1HX5l1xH76nnGy2llxC_YdtXg9_bybKexWoUXLnTM0zU' # Lubb's sheet

    # copy template
    body = {
    'name': orderNum
    }
    drive_response = DRIVE.files().copy(fileId=template_id, body=body).execute()
    presentation_copy_id = drive_response.get('id')
    print(presentation_copy_id)

    # open copied template
    presentation = SLIDES.presentations().get(presentationId=presentation_copy_id).execute()
    slides = presentation.get('slides')
    #print(slides)

    # modify copied template with wix data, starting with empty request
    requests = []
   
    # delete gpsLat data
    requests.append({
        'deleteText': {
            'objectId': 'g5b482bc06e_0_3',
            'textRange': {
                'type': 'ALL'
        }}})

    # insert gpsLat data
    requests.append({
        'insertText': {
            'objectId': 'g5b482bc06e_0_3',
            'insertionIndex': 0,
            'text': str(gpsLat)
    }})

    SLIDES.presentations().batchUpdate(presentationId=presentation_copy_id, body = {'requests': requests}).execute()

    # create request to update Google Sheet slideFlag column
    values = [
        [1]
    ]
    body = {
        'values': values
    }
    result = SHEETS.spreadsheets().values().update(spreadsheetId=sheet_id, range=slideFlagAddress, valueInputOption='USER_ENTERED',body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))

    # download as pdf
    pdf = DRIVE.files().export(fileId = presentation_copy_id, mimeType = 'application/pdf').execute()
    
    fn = '{}.pdf'.format(orderNum) #% os.path.splitext(orderNum)
    print(fn)
    with open(fn, 'wb') as fh:
        fh.write(pdf)
        print('Downloaded pdf.')
    
    # update Google Sheet pdfFlag
    values = [
        [1]
    ]
    body = {
        'values': values
    }
    result = SHEETS.spreadsheets().values().update(spreadsheetId=sheet_id, range=pdfFlagAddress, valueInputOption='USER_ENTERED',body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))


def main():
    
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    # initialize API services
    SHEETS = discovery.build('sheets', 'v4', http=http)
    SLIDES = discovery.build('slides', 'v1', http=http)
    DRIVE = discovery.build('drive', 'v3', http=http)
    #print(credentials.access_token)

    # read in wix data and convert to pd.DataFrame + clean up
    sheet_id = '1HX5l1xH76nnGy2llxC_YdtXg9_bybKexWoUXLnTM0zU' # Lubb's sheet
    wix = SHEETS.spreadsheets().values().get(range="Sheet1",spreadsheetId = sheet_id).execute().get('values')
    wix = pd.DataFrame(wix)
    wix.columns = ['firstName', 'lastName', 'gpsLat', 'gpsLong', 'orderDate', 'instagram', 'orderNum', 'slideFlag', 'pdfFlag']
    wix.drop(index=0, axis = 0, inplace=True)
    wix.reset_index(inplace=True)
    wix.drop('index', axis = 1, inplace = True)
    wix.fillna(0, inplace = True)
    print(wix)
    # add cell addresses so that Google Sheet can be updated
    wix['rowNum'] = wix.reset_index().index + 2
    wix['rowNum'] = wix['rowNum'].astype(str)
    wix['slideFlagAddress'] = 'H' + wix['rowNum']
    wix['pdfFlagAddress'] = 'I' + wix['rowNum']

    # determine which orders we need slides made and which need pdf's made
    needSlides = wix[wix['slideFlag']==0]
    needSlides.reset_index(inplace=True)
    needPdfs = wix[wix['pdfFlag']==0]
    needPdfs.reset_index(inplace=True)
    print(needSlides)
    # loop through orders that need a certificate made
    try:
        for i in range(len(needSlides)):
            
            makeCertificate(needSlides['orderNum'][i],
                            needSlides['firstName'][i],
                            needSlides['gpsLat'][i],
                            needSlides['gpsLong'][i],
                            needSlides['instagram'][i],
                            needSlides['slideFlagAddress'][i],
                            needSlides['pdfFlagAddress'][i]
                            )
    except: print("No new records.")

if __name__ == '__main__':
    main()