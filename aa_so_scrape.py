##### Importing packages
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
driver = webdriver.Chrome(ChromeDriverManager().install())
start_time = time.time()


# 1. Use BBG T-1 SO to compare with AAstock SO --> if BBG T-1 = AA SP --> BBG updated
# 2. Par value, par currency
# 2.5 board lot
# 3. voting rights
# 4. Catch totally shares --> missed domestics
# 5.
# utimate matrix to build

L ={"BBG":[1,2],
    "SSD" :[3,4,], 
    "AAstock":[5,6],
    }

u_matrix = pd.DataFrame(L, index = ["R","R-1"])

u_matrix

#####
### import ticker
hk_ticker = pd.read_excel("HK Ticker Nov_2020.xlsx")
hk_ticker.columns = ["ticker"]
stock_list = pd.DataFrame(hk_ticker.ticker.str.split(" ",1).tolist()).loc[0:,0]
stock_list

stock_list_sort = []
for i in stock_list:
    stock_list_sort.append(i)
stock_list_sort= list(map(int, stock_list_sort))
stock_list_sort.sort()
stock_list_sort

# Format ticker code into standard five digits
ticker_00000 = []
for i in stock_list:
    fn = "{num:0>5}".format(num=i)
    ticker_00000.append(fn)
ticker_00000.sort()

res_ticker = []
res_total_so = []
res_h_so = []

### Get input for how many tickers to scrape for
print(".")
print("Total " + str(len(ticker_00000)) + " tickers")
print("Enter how many tickers you want to scrape for their SO")
scrape_tic = input()
scrape_num = int(scrape_tic) # in secs
print("Scraping " + str(scrape_num) + " out of " + str(len(ticker_00000)) + " tickers")
print("--- Starting AAstock data extraction---")
print("...")

##### 
### Function for scraping so
def scrape_aa_so(i, ticker):
        driver.get('http://www.aastocks.com/en/stocks/analysis/company-fundamental/basic-information?symbol=' + ticker)
        print(i, ticker)
        issued_so = driver.find_elements_by_xpath('//*[@id="cp_pBITable1"]/div/table/tbody/tr[10]/td[2]')[0]
        h_shares_so = driver.find_elements_by_xpath('//*[@id="cp_pBITable1"]/div/table/tbody/tr[11]/td[2]')[0]
        total_so = issued_so.text
        h_so = h_shares_so.text
        if h_so == "-":
            h_so = "0"
        else:
            pass
            ###
        res_ticker.append(ticker.lstrip("0") + str(" HK Equity"))
        res_total_so.append(total_so)
        res_h_so.append(h_so)
        print(str(ticker) + str(" HK Equity")," | ", total_so," | " , h_so)

### Scraping with fail save 
for i, ticker in enumerate(ticker_00000[0:scrape_num]):
    try:
        scrape_aa_so(i,ticker)
    except IndexError:
        try:
            print(".")
            print("second attempt for " + str(i) + " position")
            print("second attempt for " + str(ticker))
            scrape_aa_so(i,ticker)
        except IndexError:
            try:
                print(".")
                print("third attempt for " + str(i) + " position")
                print("third attempt for " + str(ticker))
                scrape_aa_so(i,ticker)
            except IndexError:
                break

### return result 
print(".")
print(str((len(res_total_so)/scrape_num)*100) + "%" + " scraping succeeded" )
print("Scraped " + str(scrape_num) + " out of " + str(len(ticker_00000)) + " tickers")
print("---AAstock data extraction completed---")


##### Building export excel data format

### function 1 for BDP
def BDP(equity_list , field):
    res = []
    for i in equity_list:
            res1 = "=BDP(\"" + str(i) + " HK Equity" + "\",\"" + str(field) + "\")" 
            res.append(res1)
    
    export_list = pd.DataFrame(res)
    return export_list

### function 2 for if logic compare diff

def diff_logic(equity_list, first_col, second_col):
    res=[]
    for i, j in enumerate(equity_list):
        so_diff = "=IF(ABS(" + str(first_col) + str(i+2) + "-" + str(second_col)  + str(i+2) + ")>1,1,0)"
        res.append(so_diff)
    so_diff_list = pd.DataFrame(res)
    return so_diff_list
    
# 1. Ticker
# Ticker dataframe

stock_list_pd = pd.DataFrame(res_ticker)
                             
# 2. Fundamental Ticker
# Pull fundamental ticker with BBG API
res_fund=[]
for i in stock_list_sort:
    fund = "=RIGHT(BDP(\"" + str(i) + " HK Equity"  + "\", \"EQY_FUND_TICKER\") ,2)"
    res_fund.append(fund)
    
fund_pd = pd.DataFrame(res_fund)

# Multi Class
multi = BDP(stock_list_sort,"MULTIPLE_SHARE")

# AA Total SO
# Total_so dataframe
total_so_pd = pd.DataFrame(res_total_so)

# AA H SO
# h shares so dataframe
h_so_pd = pd.DataFrame(res_h_so)

# BBG SO total
# Bloomberg API dataframe
bbg_total_pd = BDP(stock_list_sort, "TOTAL_VOTING_SHARES_VALUE")

## BBG SO, multi only return H
bbg_h_pd = BDP(stock_list_sort, "EQY_SH_OUT_REAL")

# total_diff
total_diff = diff_logic(stock_list_sort, "E", "G") 

# H_shares_diff
h_diff = diff_logic(stock_list_sort, "F","H")

### Combine list
frame = [stock_list_pd, fund_pd, multi, total_so_pd,  h_so_pd, bbg_total_pd, bbg_h_pd, total_diff, h_diff]
result_excel = pd.concat(frame, axis = 1, sort = False)
result_excel.columns = ["Ticker","Fundamental Ticker", "Multi Class", "AA Total SO", "AA H SO", "BBG SO total",
                        "BBG SO (If multi class, only return H-shares)","Total_diff", 
                        "H_diff", ]

### Export result into excel
today = datetime.today().strftime('%Y-%m-%d')
pd.DataFrame(result_excel).to_excel("AA_SO " + str(today) +".xlsx")

print(".")
print(".")
print(".")
print("AAstock SO autocopy completed and exported excel file")
print("--- runtime: %s seconds ---" % round(time.time() - start_time))
