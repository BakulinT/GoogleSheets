from datetime import datetime, timezone
from opensea import OpenseaAPI
from opensea import utils
import json
import time
import math


class OpenSeaNFT:
    def __init__(self, key):
        self.key = key
    
    def request(self, dateTo, dateEnd, name):
        dateBase = [] 

        dateStart = [int(x) for x in dateTo.split('.')]
        dateStop = [int(x) for x in dateEnd.split('.')]

        today = datetime(dateStart[0], dateStart[1], dateStart[2], tzinfo=timezone.utc)
        yestday = datetime(dateStop[0], dateStop[1], dateStop[2], tzinfo=timezone.utc)

        secToday = today.timestamp()
        Day = yestday.toordinal() - today.toordinal()
        self.sea = OpenseaAPI(apikey=self.key)

        for d in range(0, Day + 1):
            floorPprice = float('inf')
            price = 0
            count = 1

            self.dateSelect = time.gmtime(secToday + d * 86400)
            print( time.strftime('%Y.%m.%d', self.dateSelect) )
            
            (count, price, floorPprice) = self.HourGo(count, price, floorPprice, name)

            diffFoolPrice = 0
            diffAvgPrice = 0
            lenBase = len(dateBase) - 1
            AvgStart = price / count
            
            if floorPprice == float('inf'):
                floorPprice = 0

            if ( lenBase > -1 ):
                AvgEnd = dateBase[lenBase][2] / dateBase[lenBase][3]
                
                if (AvgEnd != 0):
                    diffAvgPrice = (( AvgStart / AvgEnd ) - 1)
                if dateBase[lenBase][4] == 0:
                    diffFoolPrice = 0
                else: 
                    diffFoolPrice = (( floorPprice / dateBase[lenBase][4] ) - 1)

            dateBase.append([
                time.strftime('%Y.%m.%d', self.dateSelect), 
                AvgStart, price, count, floorPprice, diffFoolPrice, diffAvgPrice
                ])
            time.sleep(0.2)
            
        return dateBase
    
    def HourGo(self, count, price, floorPprice, name, hour=0, step=6):
        for hour in range(hour, 24, step):

            period_start = utils.datetime_utc(self.dateSelect.tm_year, self.dateSelect.tm_mon, self.dateSelect.tm_mday, hour, 0)
            period_end = utils.datetime_utc(self.dateSelect.tm_year, self.dateSelect.tm_mon, self.dateSelect.tm_mday, hour + step - 1, 59)

            try:
                result = self.sea.events(collection_slug=name,
                                    only_opensea="true", 
                                    event_type="successful",
                                    occurred_after=period_start,
                                    occurred_before=period_end,
                                    limit=300)
            except json.JSONDecodeError:
                print('JSONDecodeError')
                continue;

            except Exception as e:
                print('CustomException:', period_start, period_end, e)
                time.sleep(2)
                self.HourGo(count, price, floorPprice, name, hour, 2)
                return (count, price, floorPprice)

            if result.get( 'asset_events', False) == False:
                continue

            if step != 2 and len(result['asset_events']) == 300:
                return self.HourGo(count, price, floorPprice, name, hour, step - 2);

            for a in result['asset_events']:
                count += 1
                price += int( a['total_price'] ) / math.pow(10, 18) 

                if (price < floorPprice):
                    floorPprice = price
        return (count, price, floorPprice)
