import statistics
from openpyxl import Workbook



#Input desired gear ratios for high, mid, and low modes
def gethigh(x):
    if x >= float(8.55) and x <= float(9.45):
        return True
def getlow(x):
    if x >= float(51.3) and x <= float(56.7):
        return True
def getmid(x):
    if x >= float(29.925) and x <= float(33.075):
        return True


low= {}
mid = {}
high = {}
#input no. of teeth per gear and gear diameter
gears = {8:5,9:5.5,10:6,12:7,13:7.5,14:8,18:10,16:9,7:4.5,11:6.5,15:8.5,20:11,24:13,28:15,30:16,32:17,36:19,40:21,42:22,48:25,56:29,78:40,90:46,48:25,80:41,120:61,50:26,54:28,44:23,42:22,18:10}
gears2 = gears
gears3 = gears
gears4 = gears
gears5 = gears
gears6 = gears

#loop through every gear combination and store combinations that output the desired ratio
for i in gears.keys():
    for j in gears2.keys():
        for k in gears3.keys():
            for l in gears4.keys():
                        a = float(j/i)
                        b = float (l/k)
                        ratio = (a*b)
                        dist1 = (gears[i]/2) + (gears2[j]/2)
                        dist2 = (gears4[l]/2) + (gears3[k]/2)
                        if gethigh(ratio):
                            
                                    high[f'{i}-{j}-{k}-{l}']=[dist1, dist2]
                        if getlow(ratio):
                                    
                                    low[f'{i}-{j}-{k}-{l}']=[dist1, dist2]
                        if getmid(ratio):
                                    
                                    mid[f'{i}-{j}-{k}-{l}']=[dist1, dist2]


#open excel file and set up columns
wb = Workbook()
ws = wb.active
ws.append(['high','mid','low','dist1', 'dist2', 'hratio', 'mratio', 'lratio'])

#compare the distances of the combinations so that the distances 
#between the shaft are constant for high,mid, and low modes
for key1,val1 in high.items():
      for key2,val2 in low.items():
            for key3,val3 in mid.items():
                  if (val1 == val2) and (val2==val3):
                    hr = [int(x) for x in key1.split('-')]
                    lr = [int(x) for x in key2.split('-')]
                    mr = [int(x) for x in key3.split('-')]
                    hratio = ((hr[3]/hr[2])*(hr[1]/hr[0]))
                    lratio = ((lr[3]/lr[2])*(lr[1]/lr[0]))
                    mratio = ((mr[3]/mr[2])*(mr[1]/mr[0]))
                    x = [key1,key2,key3,val1[0],val1[1], hratio, mratio, lratio]
                    ws.append(x)
                    print("appended")

#save excel file
wb.save("gears_wratio.xlsx")
print("finished")
print(len(mid))
print(len(high))
print(len(low))