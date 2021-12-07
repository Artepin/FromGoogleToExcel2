import os.path
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os


def init(p_Creds, p_tokenName):

    # os.environ["HTTPS_PROXY"] = 'http://kozlovaa:123@prox.mera.local:3128'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    ltoken = p_tokenName+ '.json'
    if os.path.exists(ltoken):  # Не следует создавать token.json вручную!
        creds = Credentials.from_authorized_user_file(ltoken, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(p_Creds+'.json', SCOPES)
            #rename token
            #os.rename('token.json', p_tokenName)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(ltoken, 'w') as token:
            token.write(creds.to_json())

    # service = build('sheets', 'v4', credentials=creds)
    try:
        service = build('sheets', 'v4', credentials=creds)
    except:
        DISCOVERY_SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
        service = build('sheets', 'v4', credentials=creds, discoveryServiceUrl=DISCOVERY_SERVICE_URL)
    # proxies = 'https://kozlovaa:123@prox.mera.local:3128'
    ss = service.spreadsheets()
    return ss


def getCoordOfNr(p_ss,p_ssId,p_ranges):
    ss = p_ss
    ssId = p_ssId
    sheet = ss.get(spreadsheetId=ssId, ranges=p_ranges).execute()
    print(type(sheet))
    print(sheet)
    nr = sheet['namedRanges']
    score_ranges = 0
    for i in p_ranges:
        if nr[score_ranges]['name'] != i:
            for j in range(0,len(nr)):
                if str(nr[j]['name'])==i:
                    nr[j],nr[score_ranges] = nr[score_ranges],nr[j]
        score_ranges +=1
    #nr = nr[::-1]
    print(type(nr))
    print(nr)
    coords = []
    for i in range(0, len(nr)):
        print(type(nr[i]))
        print(nr[i])
        rg = nr[i]['range']
        print(type(rg))
        print(rg)
        rowId = rg['startRowIndex']
        print(type(rowId))
        print(rowId)
        columnId = rg['startColumnIndex']
        print(type(columnId))
        print(columnId)
        cellId = str(rowId+1) + ' ' + str(columnId+1)
        coords.append(cellId)
    if coords != []:
        print(coords)
        return coords
    else:
        print('An error in func getCoordOfNr')

def rowColumnCellGet(p_ss,p_ssId, p_col,p_row):
    ss = p_ss
    ssId = p_ssId
    column = str(p_col)
    row = str(p_row)
    request = ss.values().get(spreadsheetId=ssId, range='R'+row+'C'+column).execute()
    print(request)
    values = request.get('values', [])
    return values

def rowcol_to_a1(row, col):
        """Translates a row and column cell address to A1 notation.

        :param row: The row of the cell to be converted.
            Rows start at index 1.
        :type row: int, str

        :param col: The column of the cell to be converted.
            Columns start at index 1.
        :type row: int, str

        :returns: a string containing the cell's coordinates in A1 notation.

        Example:

        >>> rowcol_to_a1(1, 1)
        A1

        """
        magic_number = 64
        row = int(row)
        col = int(col)

        '''if row < 1 or col < 1:
            raise IncorrectCellLabel('(%s, %s)' % (row, col))'''

        div = col
        column_label = ''

        while div:
            (div, mod) = divmod(div, 26)
            if mod == 0:
                mod = 26
                div -= 1
            column_label = chr(mod + magic_number) + column_label

        label = '%s%s' % (column_label, row)

        return label
