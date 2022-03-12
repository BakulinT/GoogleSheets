from OpenseaApiNFT import OpenSeaNFT
import sys

keyApi = ""# Token for Opensea Api

DefaultdateStart = '2022.1.23'
dateStop = '2022.2.3'
Defaultname = "boredapeyachtclub"

args = sys.argv

largs = len(args)

if (largs == 2):
    DefaultdateStart = args[1]
elif ( largs > 2 ):
    DefaultdateStart = args[1]
    DefaultdateStop = args[2]
if (largs > 3):
    Defaultname = args[3:len(args)]

opns = OpenSeaNFT(keyApi) 

# Out: Date, Avg, Volume, Count, Floor Price, Diff Fool Price, Diff Avg Price
dateBase = opns.request(DefaultdateStart, DefaultdateStop, Defaultname)

print('\n')

p = 0
c = 0
for i in dateBase:
    p += i[1]
    c += i[2]
    print(f"[[{i[0]}], \n\tAVG: {i[1]},\n\tPrice: {i[2]}, \n\tCount: {i[3]}, \n\tFloor price: {i[4]}, \n\tDiff Floor price: {i[5]}, \n\tDiff Avg: {i[6]}")
print('\n****Total****')
try:
    print( 'Avg: {0:.3f},\nVolume: {1}\nNum. sales: {2}'.format(p / c, p, c) )
except ZeroDivisionError:
    print( 'Avg: {0:.3f},\nVolume: {1}\nNum. sales: {2}'.format(0.0, p, c) )