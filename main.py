import pandas as pd
import yfinance as yf
import datetime
from openpyxl import load_workbook
import numpy as np
end_date = datetime.datetime.date(datetime.datetime.now())
start_date = end_date - datetime.timedelta(days=3*365)
def pairtrader(a,b,c):
    data = yf.download(a+" "+b, start=start_date, end=end_date)

    data = data.drop(data.columns[[0,1,4,5,6,7,8,9]],axis=1)
    data['ratio'] = np.where(data[data.columns[1]]>data[data.columns[0]],data[data.columns[1]]/data[data.columns[0]],data[data.columns[0]]/data[data.columns[1]])

    average = data['ratio'].mean()
    sd = data['ratio'].std()
    cor = data[data.columns[1]].corr(data[data.columns[0]])
    data['Ave + 2SD'] = data['ratio'] >= average + 2*sd
    data['Ave - 2SD'] = data['ratio'] <= average - 2 * sd

    with pd.ExcelWriter(c+".xlsx", mode='a', if_sheet_exists="replace") as writer:
                data.to_excel(writer, sheet_name='Sheet1')


    workbook = load_workbook(filename = c+".xlsx")

    #open workbook
    sheet = workbook.active

    sheet["K3"] = "Average"
    sheet["K4"] = average
    sheet["I3"] = "CORRELATION"
    sheet["I4"] = cor
    sheet["L3"] = "STANDARD DEVIATON"
    sheet["L4"] = sd
    sheet["K6"] = "+1SD"
    sheet["K7"] = sd+average
    sheet["L6"] = "+2SD"
    sheet["L7"] = sd*2+average
    sheet["M6"] = "+3SD"
    sheet["M7"] = sd*3+average
    sheet["I8"] = "-1SD"
    sheet["J8"] = -sd+average
    sheet["I9"] = "-2SD"
    sheet["J9"] = -sd*2+average
    sheet["I10"] = "-3SD"
    sheet["J10"] = -sd*3+average
    workbook.save(filename = c+".xlsx")
pairtrader("KOTAKBANK.BO","HDFCBANK.BO","HDFC-KOTAK")
pairtrader("HCLTECH.NS","INFY.NS","HCL-Infosys")
pairtrader("WIPRO.NS","INFY.NS","WIPRO-Infosys")
# pairtrader("BAJFINANCE.NS","BAJAJFINSV.NS","FINSERV-FINANCE")
# pairtrader("TATASTEEL.BO","JSWSTEEL.BO","TSL-JSW")
# pairtrader("ULTRACEMCO.NS","ACC.BO","ACC-ULTRA")
# pairtrader("AMBUJACEM.NS","ACC.BO","ACC-AMBUJA")
# pairtrader("AMBUJACEM.NS","ULTRACEMCO.NS","AMBUJA-ULTRA")