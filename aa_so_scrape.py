#Importing packages
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
driver = webdriver.Chrome(ChromeDriverManager().install())
start_time = time.time()


### Chrome Scrape ticker
stock_list = ["700", "66", "823", "270", "939", "1122"]

# Format ticker code into standard five digits
res_f = []
for i in stock_list:
    fn = "{num:0>5}".format(num=i)
    res_f.append(fn)
res_f

# 
res_ticker = []
res_total_so = []
res_h_so = []

print("--- Starting AAstock data extraction---")
print("...")

for ticker in res_f:
    driver.get('http://www.aastocks.com/en/stocks/analysis/company-fundamental/basic-information?symbol=' + ticker)
    issued_so = driver.find_elements_by_xpath('//*[@id="cp_pBITable1"]/div/table/tbody/tr[10]/td[2]')[0]
    h_shares_so = driver.find_elements_by_xpath('//*[@id="cp_pBITable1"]/div/table/tbody/tr[11]/td[2]')[0]
    total_so = issued_so.text
    h_so = h_shares_so.text
    if h_so == "-":
        h_so = "0"
    else:
        pass
    ###
    res_ticker.append(ticker + str(" HK Equity"))
    res_total_so.append(total_so)
    res_h_so.append(h_so)
    
    
print("---AAstock data extraction completed---")

# Ticker dataframe
 
stock_list_pd = pd.DataFrame(res_ticker)
                             
# Pull fundamental ticker with BBG API
res_fund=[]
for i in stock_list:
    fund = "=RIGHT(BDP(B2, \"EQY_FUND_TICKER\") ,2)"
    res_fund.append(fund)
    
fund_pd = pd.DataFrame(res_fund)

# Total_so dataframe
total_so_pd = pd.DataFrame(res_total_so)

# h shares so dataframe
h_so_pd = pd.DataFrame(res_h_so)

# Bloomberg API dataframe
res_api=[]
for i in stock_list:
    API = "=BDP(\"" + str(i) + "\", " "\"EQY_SH_OUT_REAL\")"
    res_api.append(API)

API_pd = pd.DataFrame(res_api)

# Excel filter function for SO different between BBG and AA stock > 1
res_diff=[]
for i in stock_list:
    filter_1 = "=if(ABS(C2-E2)>1 ,1 ,0)"
    res_diff.append(filter_1)
    
filter_pd = pd.DataFrame(res_diff)

# Combine list
frame = [stock_list_pd, fund_pd, total_so_pd,  h_so_pd, API_pd, filter_pd]
result_excel = pd.concat(frame, axis = 1, sort = False)
result_excel.columns = ["Ticker","Fundamental Ticker", "AA Total SO", "AA H SO", "BBG SO", "filter_pd"]

# Export result into excel
today = datetime.today().strftime('%Y-%m-%d')
pd.DataFrame(result_excel).to_excel("AA_SO " + str(today) +".xlsx")

print("---")
print("AAstock SO autocopy completed and exported excel file")
print("--- runtime: %s seconds ---" % round(time.time() - start_time))
