from googleSheetsApi import SheetsApi

IDSheets = '' # Id table Sheets
keyOpenseApi = '' # Key for Opensea Api

sht = SheetsApi(IDSheets, keyOpenseApi)

sht.ConectToSheets() 
# sht.NewDate()
# sht.Date()
# sht.AddToEnd()
sht.ControlCollections()
