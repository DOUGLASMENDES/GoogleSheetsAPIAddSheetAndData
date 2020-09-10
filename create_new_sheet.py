# Author: Douglas Mendes


from Google import Create_Service
import pandas as pd
from numpy.random import randint

# Nome da aba a ser criada:
sheetName = 'NEW-SHEET-NAME'

# Id da planilha do Google Sheets:
gsheet_id = '<GOOGLE SHEET FILE ID>'  # ex: 'acrEdwekci89tLkeT_e605K2AET6ug7XCVIerXC2YjBs'

# Arquivo JSon com Client ID autorizando escrita no Google Sheets:
CLIENT_SECRET_FILE = '<PATH TO CLIENT SECRET JSON FILE>'

# DataFrame de exemplo:
df = pd.DataFrame(columns=['lib', 'qty1', 'qty2'])
for i in range(200):
    df.loc[i] = ['name' + str(i)] + list(randint(10, size=2))


API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)
          
spreadsheets = service.spreadsheets()

def add_sheets(gsheet_id, sheet_name):
    try:
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                        'tabColor': {
                            'red': 0.44,
                            'green': 0.99,
                            'blue': 0.50
                        }
                    }
                }
            }]
        }

        response = spreadsheets.batchUpdate(
            spreadsheetId=gsheet_id,
            body=request_body
        ).execute()

        return response
    except Exception as e:
        print(e)


def InputaDataframeNovaAba(df, sheetName, gsheet_id):
    # Transformando o dataframe em uma lista de listas
    values = []
    # Linha do título
    values.append(list(df.columns))

    # Adicionando linha por linha:
    for i in range (0,len(df)):
        values.append(list(df.iloc[i]))

    sheetRange = sheetName + "!A1:C50000"

    batch_update_values_request_body = {

        # How the input data should be interpreted.
        'value_input_option': 'USER_ENTERED',  

        # The new values to apply to the spreadsheet.
        "data": [
        {
        "range": "NEW-SHEET-NAME!A1:C50000",
        "values": values,
        "majorDimension": "ROWS"
        }
    ]
    }

    request = service.spreadsheets().values().batchUpdate(spreadsheetId=gsheet_id, body=batch_update_values_request_body)
    response = request.execute()



# Adiciona novas abas (sheets):
xSheets = [ sheetName ] # , 'nova-aba2', 'nova-aba3', 'nova-aba4']
for name in xSheets:
    print(add_sheets(gsheet_id, name))

# Coloca os dados na nova aba:
InputaDataframeNovaAba(df, sheetName, gsheet_id)

