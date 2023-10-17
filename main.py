# Imports
from yahoofinancials import YahooFinancials
import numpy as np
from datetime import date
from datetime import timedelta
import etflist as etflist
import xlsxwriter
from statistics import mean
print("This Software is the Intellectual Property of Brandon M. Miller. ")
print("It May NOT Be Used Under ANY Circumstance Without the EXPRESS ")
print("WRITTEN PERMISSION of the author. This Software Also Contains ")
print("Open-Source Libraries Not Covered by This Statement ")
print("These Libraries are NOT Property of Brandon M. Miller.")
print("Credit Given to John Johnson for the Creation of the Algorithm that Allows this Program to Function")
rlist = []
bfslist = []
# Create Excel File
outWorkbook = xlsxwriter.Workbook("Output.xlsx")
outSheet = outWorkbook.add_worksheet()
outSheet.write("A1", "Ticker")
outSheet.write("B1", "R-Value")
outSheet.write("C1", "Slope")
# Write Tickers
c = 1
while c == 1:
    for item in range(len(etflist.etflist)):
        outSheet.write(c, 0, etflist.etflist[item])
        c += 1
# Date Calc
# XLE
today = date.today()
startdate = date.today() - timedelta(180)
todaystring = str(today)
startdatestring = str(startdate)
z = 0
while z < len(etflist.etflist):
    # Stock Prices
    ticker = str(etflist.etflist[z])
    yahoo_financials = YahooFinancials(ticker)
    pricing = yahoo_financials.get_historical_price_data(startdatestring, todaystring, 'daily')
    pricelist = []
    n = 65
    for i in range(0, n):
        p = (pricing[ticker]['prices'][i]['close'])
        pricelist.append(p)

    # Calc R value

    x_values = range(0, 65)
    y_values = [pricelist]
    correlation_matrix = np.corrcoef(x_values, y_values)
    correlation_xy = correlation_matrix[0, 1]
    print(ticker)
    print('R Value:')
    print(correlation_xy)
    rlist.append(correlation_xy)
    correlation_xy = 0

    # Calc Best Fit Slope

    xs = np.array(range(0, 65), dtype=np.float64)
    ys = np.array(pricelist, dtype=np.float64)


    def best_fit_slope(xs, ys):
        m = (((mean(xs) * mean(ys)) - mean(xs * ys)) /
             ((mean(xs) ** 2) - mean(xs ** 2)))
        return m


    m = best_fit_slope(xs, ys)
    print("Best Fit Slope:")
    print(m)
    bfslist.append(m)
    m = 0

    if z == len(etflist.etflist)-1:
        c = 1
        for item in range(len(rlist)):
            outSheet.write(c, 1, rlist[item])
            c += 1
        d = 1
        for item in range(len(bfslist)):
            outSheet.write(d, 2, bfslist[item])
            d += 1
        outWorkbook.close()

    if z <= len(etflist.etflist)-1:
        z += 1
