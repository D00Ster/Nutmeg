import urllib.request, http.client, json, xlwt
from datetime import datetime

font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.colour_index = 2
font0.bold = True
style1 = xlwt.XFStyle()
wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')


def getResults(start, monthly, risk, time):
    data = "?investment_style=managed&lump_sum={}&monthly={}&risk_level={}&timeframe={}&type=standard".format(start, monthly, risk, time)
    investments = urllib.request.urlopen("https://app.nutmeg.com/client/funds/chart/investments" + data)
    projection = urllib.request.urlopen("https://app.nutmeg.com/client/funds/chart/projection" + data)
   
    return investments.read().decode(), projection.read().decode()

TIME = [5, 50]
MONTHLY = [2000]
START = [25000]
RISK = [1, 2, 3, 4, 5, 6, 7, 8, 9]


ws.write(0, 0, 'Time', style1)
ws.write(0, 1, 'Start', style1)
ws.write(0, 2, 'Monthly', style1)
ws.write(0, 3, 'Risk', style1)

ws.write(0, 4, 'Expected in 12 months', style1)
ws.write(0, 5, 'Expectes final', style1)

ws.write(0, 6, 'Portfolio gain 12 months', style1)
ws.write(0, 7, 'Portfolio gain final', style1)

ws.write(0, 8, 'Portfolio gain % 12 months', style1)
ws.write(0, 9, 'Portfolio gain % final', style1)

ws.write(0, 10, 'Annual costs 12 months', style1)
ws.write(0, 11, 'Annual costs final', style1)

ws.write(0, 12, 'Annual costs % 12 months', style1)
ws.write(0, 13, 'Annual costs % final', style1)

ws.write(0, 14, 'Cumulative costs 12 months', style1)
ws.write(0, 15, 'Cumulative costs final', style1)

ws.write(0, 16, 'Cumulative costs % 12 months', style1)       
ws.write(0, 17, 'Cumulative costs % final', style1)
         
ws.write(0, 18, 'Tickers', style1)


counter = 1
print('time, start, monthly, risk')
for time in TIME:
    for start in START:
        for monthly in MONTHLY:
            for risk in RISK:
                investments, projection= getResults(start, monthly, risk, time)
                print(time, start, monthly, risk, end=' ')
               
                a = json.loads(projection)
                investments = json.loads(investments)
                ws.write(counter, 0, time)
                ws.write(counter, 1, start)
                ws.write(counter, 2, monthly)
                ws.write(counter, 3, risk)
               
                print(a['series']['P50']['expectedReturns'][12], end=' ') # expected in 12 months
                print(a['series']['P50']['expectedReturns'][-1], end=' ') # expected final
                ws.write(counter, 4, a['series']['P50']['expectedReturns'][12])
                ws.write(counter, 5, a['series']['P50']['expectedReturns'][-1])
               
                print(a['series']['P50']['gains'][12], end=' ') # portfolio gain 12 months
                print(a['series']['P50']['gains'][-1], end=' ') # portfolio gain final
                ws.write(counter, 6, a['series']['P50']['gains'][12])
                ws.write(counter, 7, a['series']['P50']['gains'][-1])
                
                print(round((a['series']['P50']['gainsPercentage'][12]) * 100, 2), end=' ') # portfolio gain % 12 months
                print(round((a['series']['P50']['gainsPercentage'][-1]) * 100, 2), end=' ') # portfolio gain % final
                ws.write(counter, 8, a['series']['P50']['gainsPercentage'][12])
                ws.write(counter, 9, a['series']['P50']['gainsPercentage'][-1])
       
                print(a['series']['P50']['annualCosts'][12], end=' ') # annual costs 12 months
                print(a['series']['P50']['annualCosts'][-1], end=' ') # annual costs final
                ws.write(counter, 10, a['series']['P50']['annualCosts'][12])
                ws.write(counter, 11, a['series']['P50']['annualCosts'][-1])
               
                print(round(a['series']['P50']['annualCostsPercentage'][12] * 100, 2), end=' ') # annual costs % 12 months
                print(round(a['series']['P50']['annualCostsPercentage'][-1] * 100, 2), end=' ') # annual costs % final
                ws.write(counter, 12, round(a['series']['P50']['annualCostsPercentage'][12] * 100, 2))
                ws.write(counter, 13, round(a['series']['P50']['annualCostsPercentage'][-1] * 100, 2))
                
                print(a['series']['P50']['cumulativeCosts'][12], end=' ') # cumulative costs 12 months
                print(a['series']['P50']['cumulativeCosts'][-1], end=' ') # cumulative costs final
                ws.write(counter, 14, a['series']['P50']['cumulativeCosts'][12])
                ws.write(counter, 15, a['series']['P50']['cumulativeCosts'][-1])
                
                print(round(a['series']['P50']['cumulativeCostsPercentage'][12] * 100, 2), end=' ') #  cumulative costs % 12 months
                print(round(a['series']['P50']['cumulativeCostsPercentage'][-1] * 100, 2), end=' ') #  cumulative costs % final
                ws.write(counter, 16, round(a['series']['P50']['cumulativeCostsPercentage'][12] * 100, 2))
                ws.write(counter, 17, round(a['series']['P50']['cumulativeCostsPercentage'][12] * 100, 2))
                b = []
                for i in range(len(investments)):
                    print(investments[i - 1]['code'], investments[i - 1]['allocation'], end=' ') # tickers
                    b.append(investments[i - 1]['code'])
                    b.append(' ')
                    b.append(investments[i - 1]['allocation'])
                    b.append(' ')
                print(' ')
                ws.write(counter, 18, b)
                counter += 1
wb.save('NutmegABC.xls')