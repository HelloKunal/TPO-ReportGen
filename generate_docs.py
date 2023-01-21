from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from get_data_from_sheets import GetValues

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/documents.readonly', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/drive.file']

# The ID of a sample document.
TEMPLATE_DOCUMENT_ID = '1_VZmnePHNcu9ZWwPV3chEVfvYgXlzKUnYu8YJk7viTo'
DOCUMENT_ID = '1r1R4xHIaioW-A1oykNp2WhsjNpY57OknUIoI4M0z_lo'

class GenerateDocs:

    def __init__(self) -> None:
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        self.creds = creds

    def createNew(self):
        COPY_TITLE = 'Copy Title'
        BODY = {
            'name': COPY_TITLE
        }
        service = build('drive', 'v3', credentials=self.creds)
        drive_response = service.files().copy(
            fileId=TEMPLATE_DOCUMENT_ID, body=BODY).execute()
        document_copy_id = drive_response.get('id')
        print(document_copy_id)

    
    def generateTable(self, text):
        try:
            service = build('docs', 'v1', credentials=self.creds)

            # Retrieve the documents contents from the Docs service.
            document = service.documents().get(documentId=DOCUMENT_ID).execute()

            print('The title of the document is: {}'.format(document.get('title')))
              # Insert a table at the end of the body.
  # (An empty or unspecified segmentId field indicates the document's body.)
            
            row = len(text)
            column = len(text[0])


            requests = []   
            insert_table_data = {
                'insertTable': {
                    'rows': row,
                    'columns': column,
                    'location': {'index': 1}
                },
            }         
            requests.append(insert_table_data)

            table_data = []
            index = 0
            for i in reversed(range(row)):
                inp_text = text[i]
                rowIndex = (i+1)*5
                for j in reversed(range(column)):
                    index = rowIndex + (j*2)
                    insert_value = {
                        "insertText":
                            {
                                "text": str(inp_text[j]),
                                "location":
                                    {
                                        "index": index
                                    }
                            }

                    }
                    print(insert_value)
                    table_data.append(insert_value)

            # print(table_data)
            requests.append(table_data)
            result = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()

        except HttpError as err:
            print(err)


if __name__ == '__main__':
    get_values_object = GetValues()
    data = get_values_object.compl_proc()
    generate_docs_object = GenerateDocs()
    # generate_docs_object.createNew()
    fte_data = data[0]
    fte_data = fte_data[:10]
    # print(fte_data)
    generate_docs_object.generateTable(fte_data)