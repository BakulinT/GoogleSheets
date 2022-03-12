from oauth2client.service_account import ServiceAccountCredentials
from OpenseaApiNFT import OpenSeaNFT
import apiclient.discovery
import datetime
import time
import math
import httplib2

class SheetsApi:
    def __init__(self, id, keyOpensea):
        self.spreadsheet_id = id # id table sheets from url
        self.requests = []
        self.NumbColum = 999
        self.NumbRows = 2988
        self.apiKeySheets = keyOpensea

    def ConectToSheets(self, file=r"credentials.json"):
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
            file, # SFile heets Google Api
            ['https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'])
        httpAuth = credentials.authorize(httplib2.Http())
        
        self.service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

    def ConvertToBase(self, num):
        alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        b = alpha[num % 26] 
        while num >= 26 :
            num = num // 26
            b += alpha[num % 26 - 1] 
        return b[::-1] 
    
    def RoundBase(self, dateBase, x=-1, y=-1):
        AvgEnd = 0
        floorEnd = 0
        if x != -1:
            lColumStart = self.ConvertToBase( x )
            values = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"CollectionsBase!{lColumStart}{y + 3}:{lColumStart}{y + 8}",
                majorDimension='ROWS'
            ).execute()

            if values.get( 'values', False) == False:
                AvgEnd = 0
            else:
                floorEnd = float( values['values'][4][0].replace(',', '.').replace('%', '') )
                
                try:
                    AvgEnd = float( values['values'][2][0].replace(',', '.') ) / float( values['values'][3][0].replace(',', '.') )
                except ZeroDivisionError:
                    AvgEnd = 0
        for i in dateBase:
            i[0] = i[0]             # date
            i[1] = round(i[1], 2)   # AvgStart
            i[2] = round(i[2])      # price
            i[3] = i[3]             # count
            i[4] = round(i[4], 2)   # floorPprice
            if floorEnd != 0:
                i[5] = round( (i[4] / floorEnd) - 1, 3)   # diffFoolPrice
            else:
                i[5] = 1
            if AvgEnd != 0:
                i[6] = round( (i[1] /  AvgEnd) - 1, 3)   # diffAvgPrice
            else:
                i[6] = 1
            floorEnd = i[4]
            AvgEnd = i[1]

        return dateBase
    # 
    def CountTableRows(self, nameList):
        values = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=f"{nameList}!B2:B{self.NumbColum}",
            majorDimension='ROWS'
        ).execute()
        if values.get( 'values', False) == False:
            return 2

        for i in range(0, len(values['values']), 8):
            if len(values['values'][i]) == 0:
                return i + 2
        return len(values['values']) + 9

    # Returns a list of sheets
    def VeiwList(self):
        sheetsList = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute().get('sheets')
        return sheetsList

    def ListCollections(self, nameList='ListCollections', range='A3:D', major='COLUMNS'):
        values = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=f"{nameList}!{range}{self.NumbRows}",
            majorDimension=major # COLUMNS
        ).execute()

        if values.get( 'values', False) == False:
            return  [[]]
        return values['values']


    def addCollection(self, nameList, name, dateBase, it):
        lColum = self.ConvertToBase( len(dateBase) + 3 )
        listId = 0
        for i in self.VeiwList():
            if i['properties']['title'] == nameList:
                listId = i['properties']['sheetId']
        values = self.service.spreadsheets().values().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={ 
                "valueInputOption": "USER_ENTERED",
                "data": [
                    {
                        "range": f"{nameList}!B{it}:B{it + 6}",
                        "majorDimension": "COLUMNS",
                        "values": [[name]]
                    },
                    {
                        "range": f"{nameList}!C{it}:C{it + 6}",
                        "majorDimension": "COLUMNS",
                        "values": [['Data', 'Avg', 'Volume', 'Count', 'Floor price', 'Diff Floor price', 'Diff Avg price']]
                    },
                    {
                        "range": f"{nameList}!D{it}:{lColum}{it + 6}",
                        "majorDimension": "COLUMNS",
                        "values": dateBase
                    }
                ]
            }
        ).execute()
        self.requests.append(
            {
                'mergeCells': {'range': {'sheetId': listId,
                    'startRowIndex':  it - 1, # 2 - 1
                    'endRowIndex': it + 6,
                    'startColumnIndex': 1, # C
                    'endColumnIndex': 2}, # D
                    'mergeType': 'MERGE_ALL'}
            })
        self.requests.append(
            {
                'repeatCell': {'range': {'sheetId': listId,
                    'startRowIndex': it - 1,
                    'endRowIndex': it + 6,
                    'startColumnIndex': 1,
                    'endColumnIndex': 2
                },
                'cell': {'userEnteredFormat': {
                                'horizontalAlignment': 'CENTER',
                                'verticalAlignment': 'MIDDLE',
                                'numberFormat': {'type': 'Text'},
                                'textFormat': {'bold': True}
                            }
                    },
                'fields': 'userEnteredFormat'}
            })
        result = self.service.spreadsheets().batchUpdate(
            spreadsheetId = self.spreadsheet_id,
            body={
                "requests": self.requests
            }
        ).execute()

    def Availabilitycheck(self, listCollections):
        dell = []
        count = 2
        values = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=f"CollectionsBase!B2:B{self.NumbColum}",
            majorDimension='COLUMNS'
        ).execute()
        if values.get( 'values', False) == False:
            return []
        for i in range(0, len(values['values']), 8):
            if values['values'][0][i] != '' and not values['values'][0][i] in listCollections[0]:
                dell.append([values['values'][i], count])
            count += 1
        return dell

    # Dell table rows and archiving
    def CollectionsArchive(self, dell, listCollectionsBase, check=False):
        if len(dell) == 0:
            return
        lColum = self.ConvertToBase( self.NumbColum )
        for d in dell:
            dateBase = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"CollectionsBase!C{d[1]}:{lColum}{d[1] + 6}",
                majorDimension='COLUMNS'
            ).execute()['values']

            clear = False
            it = self.CountTableRows('CollectionsBaseArchive')

            listArchiv = self.ListCollections('CollectionsBaseArchive', 'B2:B')[0]
            if (len(listArchiv) + 8) > self.NumbRows:
                self.ListAdd('CollectionsBaseArchive')
                            
            for i in range(0, len(listArchiv)):
                if d[0] == listArchiv[i]:
                    clear = True
                    it = i + 2
            self.UpdateRecords(d[0], dateBase[1:len(dateBase)], it, clear=clear)

            self.service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=f"CollectionsBase!B{d[1]}:{lColum}{d[1] + 6}" # 
            ).execute()
            listCollectionsBase[d[1] - 2] = ''
        if check:
            return
        for d in dell:
            for i in range(len(listCollectionsBase) - 1, d[1] + 5, -8):
                if listCollectionsBase[i] != '':
                    dateBase = self.service.spreadsheets().values().get(
                        spreadsheetId=self.spreadsheet_id,
                        range=f"CollectionsBase!C{i + 2}:{lColum}{i + 8}",
                        majorDimension='COLUMNS'
                    ).execute()['values']

                    self.addCollection('CollectionsBase', listCollectionsBase[i], dateBase[1:len(dateBase)], d[1])
                    
                    listCollectionsBase[d[1] - 2] = listCollectionsBase[i]

                    self.service.spreadsheets().values().clear(
                        spreadsheetId=self.spreadsheet_id,
                        range=f"CollectionsBase!B{i + 2}:{lColum}{i + 7}"
                    ).execute()
                    listCollectionsBase[i] = ''
                    break
    
    def ListAddNew(self, nameList):
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={
                "requests": [
                        {
                        "addSheet": {
                            "properties": {
                                "title": nameList,
                                "gridProperties": {
                                    "rowCount": self.NumbColum,
                                    "columnCount": 999
                                    }
                                }
                            }
                        }
                    ]
            }
        ).execute()

    def ListAdd(self, nameList='CollectionsBase'):
        listView = self.VeiwList()
        listId = 0
        for i in self.VeiwList():
            if i['properties']['title'] == nameList:
                listId = i['properties']['sheetId']
        dateTime = datetime.datetime.now().strftime("%Y.%m.%d:%H.%M")

        for list in listView:
            if list['properties']['title'] == nameList:
                listId = list['properties']['sheetId']
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={
                "requests": [{
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": listId,
                            "title": dateTime,
                            "gridProperties":{
                                "rowCount": self.NumbRows,
                                "columnCount": self.NumbColum
                            }
                        },
                        "fields": "title"
                        }
                    }]
                }
        ).execute()

        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={
                "requests": [
                        {
                        "addSheet": {
                            "properties": {
                                "title": nameList,
                                "gridProperties": {
                                    "rowCount": 2000,
                                    "columnCount": 999
                                    }
                                }
                            }
                        }
                    ]
            }
        ).execute()
        self.TabelFormat(nameList)

    def UpdateRecords(self, name, dateBase, it, nameList='CollectionsBaseArchive', clear=False):
        lColum = self.ConvertToBase( self.NumbColum + 4 )

        if clear:
            self.service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=f"CollectionsBaseArchive!C{it}:{lColum}{it + 8}" # 6
            ).execute()

        self.addCollection(nameList, name, dateBase, it)
    
    def ChangeDate(self, listCollections):
        lColum = self.ConvertToBase( self.NumbColum )
        value = self.ListCollections(f'CollectionsBase', f'B2:{lColum}')
        if len(value[0]) == 0:
            return

        lenVal = len(value[0])
        for y in range(0, lenVal, 8):
            for i in range(0, len(listCollections[0]), 1):
                tempDateStart = listCollections[2][i].split('.')
                tempDateEnd = listCollections[3][i].split('.')
                collDateStart = datetime.date( int(tempDateStart[0]), int(tempDateStart[1]), int(tempDateStart[2]) ).strftime("%Y.%m.%d")
                collDateEnd = datetime.date( int(tempDateEnd[0]), int(tempDateEnd[1]), int(tempDateEnd[2]) ).strftime("%Y.%m.%d")

                for x in range(2, len(value), 1):
                    if y > len(value[x]) - 1 or  value[x][y] == '':
                        break
                    if x >= len(value) - 1 or ( len(value[x]) > len(value[x + 1]) and y >= len(value[x + 1]) + 1  ) or value[x + 1][y] == '':
                        if listCollections[0][i] == value[0][y] and ( value[2][y] != collDateStart or value[x][y] != collDateEnd ):
                            self.service.spreadsheets().values().clear(
                                spreadsheetId=self.spreadsheet_id,
                                range=f"CollectionsBase!B{y + 2}:{lColum}{y + 8}" # 6
                            ).execute()
                            break

    def TabelFormat(self, nameList):
        requests = []
        for i in self.VeiwList():
            if i['properties']['title'] == nameList:
                sheetId = i['properties']['sheetId']

        for i in range(0, math.trunc(self.NumbRows / 8) ):
            requests.append(
                {
                    'repeatCell': {'range': {'sheetId': sheetId,
                        'startRowIndex': 6 + 8*i,
                        'endRowIndex': 8+ 8*i,
                        'startColumnIndex': 3,
                        'endColumnIndex': self.NumbColum
                    },
                    'cell': {'userEnteredFormat': {
                                'numberFormat': {'type': 'PERCENT'}
                            }
                        },
                    'fields': 'userEnteredFormat'}
                })
            requests.append([
                    {
                    "addConditionalFormatRule": {
                        "rule": {
                        "ranges": [
                            {
                            "sheetId": sheetId,
                            'startRowIndex':  2 + + 8*i, # 2 - 1
                            'endRowIndex': 3 + + 8*i,
                            'startColumnIndex': 3, # C
                            'endColumnIndex': self.NumbColum
                            }
                        ],
                        "gradientRule": {
                            "minpoint": {
                            "color": {
                                    "green": 1,
                                    "red": 1,
                                    "blue": 1
                            },
                            "type": "MIN"
                            },
                            "maxpoint": {
                            "color": {
                                    "green": 0.64,
                                    "red": 0.34,
                                    "blue": 0.35
                            },
                            "type": "MAX"
                            },
                        }
                        },
                        "index": 0
                        }
                    }, 
                    {
                    "addConditionalFormatRule": {
                        "rule": {
                        "ranges": [
                            {
                            "sheetId": sheetId,
                            'startRowIndex':  3 + + 8*i, # 2 - 1
                            'endRowIndex': 4 + + 8*i,
                            'startColumnIndex': 3, # C
                            'endColumnIndex': self.NumbColum
                            }
                        ],
                        "gradientRule": {
                            "minpoint": {
                            "color": {
                                    "green": 1,
                                    "red": 1,
                                    "blue": 1
                            },
                            "type": "MIN"
                            },
                            "maxpoint": {
                            "color": {
                                    "green": 0.64,
                                    "red": 0.34,
                                    "blue": 0.35
                            },
                            "type": "MAX"
                            },
                        }
                        },
                        "index": 0
                        }
                    },
                    {
                    "addConditionalFormatRule": {
                        "rule": {
                        "ranges": [
                            {
                            "sheetId": sheetId,
                            'startRowIndex':  4 + + 8*i, # 2 - 1
                            'endRowIndex': 5 + + 8*i,
                            'startColumnIndex': 3, # C
                            'endColumnIndex': self.NumbColum
                            }
                        ],
                        "gradientRule": {
                            "minpoint": {
                            "color": {
                                    "green": 1,
                                    "red": 1,
                                    "blue": 1
                            },
                            "type": "MIN"
                            },
                            "maxpoint": {
                            "color": {
                                    "green": 0.64,
                                    "red": 0.34,
                                    "blue": 0.35
                            },
                            "type": "MAX"
                            },
                        }
                        },
                        "index": 0
                        }
                    },
                    {
                    "addConditionalFormatRule": {
                        "rule": {
                        "ranges": [
                            {
                            "sheetId": sheetId,
                            'startRowIndex':  5 + + 8*i,
                            'endRowIndex': 6 + + 8*i,
                            'startColumnIndex': 3,
                            'endColumnIndex': self.NumbColum
                            }
                        ],
                        "gradientRule": {
                            "minpoint": {
                            "color": {
                                    "green": 1,
                                    "red": 1,
                                    "blue": 1
                            },
                            "type": "MIN"
                            },
                            "maxpoint": {
                            "color": {
                                    "green": 0.64,
                                    "red": 0.34,
                                    "blue": 0.35
                            },
                            "type": "MAX"
                            },
                        }
                        },
                        "index": 0
                        }
                    }
            ])
        result = self.service.spreadsheets().batchUpdate(
            spreadsheetId = self.spreadsheet_id,
            body={
                "requests": requests
            }

            ).execute()
    
    def NewDate(self):
        nameList = 'CollectionsBase'
        date = datetime.date.today();
        self.service.spreadsheets().values().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={ 
                "valueInputOption": "USER_ENTERED",
                "data": [
                    {
                        "range": f"CollectionsBase!B1:B1",
                        "majorDimension": "COLUMNS",
                        "values": [[ date.strftime('%Y.%m.%d') ]]
                    }
                ]
            }
        ).execute()

    def AddToEnd(self):
        listCollections = self.ListCollections()
        listCollectionsBase = self.ListCollections('CollectionsBase', 'B2:B')[0]

        self.CheckCollections(listCollections, listCollectionsBase)      

        nameList='CollectionsBase'
        listView = self.VeiwList()
        listId = 0

        for list in listView:
            if list['properties']['title'] == nameList:
                listId = list['properties']['sheetId']

        lColum = self.ConvertToBase( self.NumbColum )
        value = self.ListCollections(f'CollectionsBase', f'B2:{lColum}')
        if len(value[0]) == 0:
            return
        lenVal = len(value[0])
        for y in range(0, lenVal, 8):
            for x in range(2, len(value), 1):
                if y > len(value[x]) - 1 or  value[x][y] == '': #  or y > len(value[x + 1]) - 1
                    break
                if x >= len(value) - 1 or ( len(value[x]) > len(value[x + 1]) and y >= len(value[x + 1]) + 1  ) or value[x + 1][y] == '':
                    lColumStart = self.ConvertToBase( x + 2 )
                    tempDateEnd = value[x][y].split('.')
                    dateNow = datetime.date.today()
                    dateEnd = datetime.date( int(tempDateEnd[0]), int(tempDateEnd[1]), int(tempDateEnd[2]) )

                    today = datetime.datetime(dateEnd.year, dateEnd.month, dateEnd.day, tzinfo=datetime.timezone.utc)
                    secToday = today.timestamp()
                    day = dateNow.toordinal() - dateEnd.toordinal()

                    if day < 2:
                        break
                    opns = OpenSeaNFT(self.apiKeySheets)
                    dateBase = self.RoundBase( opns.request( 
                        time.strftime('%Y.%m.%d', time.gmtime(secToday + 86400)),
                        time.strftime('%Y.%m.%d', time.gmtime(secToday + (day - 1) * 86400)),
                        value[0][y] ), x + 1, y)

                    values = self.service.spreadsheets().values().batchUpdate(
                                spreadsheetId=self.spreadsheet_id,
                                body={ 
                                    "valueInputOption": "USER_ENTERED",
                                    "data": [
                                        {
                                            "range": f"CollectionsBase!{lColumStart}{y + 2}:{lColum}{y + 8}", # Range?
                                            "majorDimension": "COLUMNS",
                                            "values": dateBase
                                        }
                                    ]
                                }
                            ).execute()
                    break

    # Data recording
    def Date(self):
        listCollections = self.ListCollections()
        self.ChangeDate(listCollections)
        listCollectionsBase = self.ListCollections('CollectionsBase', 'B2:B')[0]

        self.CheckCollections(listCollections, listCollectionsBase)

    def CheckCollections(self, listCollections, listCollectionsBase):
        self.CollectionsArchive( 
            self.Availabilitycheck(listCollections), 
            listCollectionsBase)
        opns = OpenSeaNFT(self.apiKeySheets) # Token for Opensea Api
        for i in range(0, len(listCollections[0])):
            dateBase = []
            if not listCollections[0][i] in listCollectionsBase: 
                time.sleep(3)
                # Out: Date, Avg, Volume, Count, Floor Price, Diff Fool Price, Diff Avg Price
                dateBase = self.RoundBase( opns.request(listCollections[2][i], listCollections[3][i],  listCollections[0][i]) )
                try:
                    it = self.CountTableRows('CollectionsBase')
                    self.addCollection('CollectionsBase', listCollections[0][i], dateBase, it) 
                except ConnectionResetError:
                    self.ConectToSheets()
                    it = self.CountTableRows('CollectionsBase')
                    self.addCollection('CollectionsBase', listCollections[0][i], dateBase, it)

    def ControlCollections(self):
        listCollections = self.ListCollections()
        listCollectionsBase = self.ListCollections('CollectionsBase', 'B2:B')[0]
        
        self.CheckCollections(listCollections, listCollectionsBase)