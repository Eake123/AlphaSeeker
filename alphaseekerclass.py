import urllib.request
from xlutils.copy import copy
from openpyxl import Workbook
from time import time as timer
from bs4 import BeautifulSoup as html
import datetime
import pandas as pd
import numpy as np
from numpy import array, shape, newaxis, isnan, arange, zeros, dot, linspace
from scipy.optimize import minimize
import random
import concurrent.futures
from multiprocessing import Pool, cpu_count
from sklearn.linear_model import LinearRegression
from fake_useragent import UserAgent

# This python script creates the stock object that is used to perform linear regression to assess the current
# value of a company based on a few financials to how it was previously valued in the past
# adj = adjacent close: The adjusted closing price amends a stock's closing price to reflect that 
# stock's value after accounting for any corporate actions. It is often used when examining historical returns or doing a detailed analysis of past performance.
# https://www.investopedia.com/terms/a/adjusted_closing_price.asp

# Stock Class
class Stock:
    def __init__(self,symbol):
        # symbol = ticker
        self.symbol = symbol
    def get_predicted(self):
        start = timer()
        # urls used to scrape certain info
        urlCash = "https://finance.yahoo.com/quote/" + self.symbol + "/cash-flow?p=" + self.symbol
        urlIncome = "https://finance.yahoo.com/quote/" + self.symbol + "/financials?p=" + self.symbol
        urlADJ = "https://ca.finance.yahoo.com/quote/" + self.symbol + "/history?period1=1462406400&period2=1620172800&interval=1mo&filter=history&frequency=1mo&includeAdjustedClose=true"
        
        
        # finds the 5 year monthly ADJ price. Reason why this isn't multithreaded is because the regression only works if there
        # is roughly 24 months (changes depending on when they release their financial statements) so if there isn't enough info
        # there's no point in sending two extra requests
        soupADJ = searchWeb(urlADJ)
        ADJList = ADJParse(soupADJ)
        urls = [urlCash, urlIncome]
        if(len(ADJList) >= 24):
            # multithreading for financial info
            with concurrent.futures.ThreadPoolExecutor() as executor:
                future = [executor.submit(searchWeb, url) for url in urls]
                soupList = [f.result() for f in future]
            freeList, issuanceDebt, endCashPosition, dateList = cashParse(soupList[0])
            revenueList, EBIT = incomeParse(soupList[1])

            # sends the info to do the math
            predicted, y_predict, currentReturn = cashMath(freeList, issuanceDebt, endCashPosition, dateList, revenueList, EBIT, ADJList)
            print("---------------------------------%s took %f seconds to complete---------------------------------" %(self.symbol,timer()-start))
            return predicted, y_predict, currentReturn
        else:
            # returns false so it's easy to know if the program should use the info or not
            return False, False, False

# this is where I use beautiful soup to begin the scraping
def searchWeb(url):
    # when finding adjacent close price I use lxml while a statement I use html.parser so this if statement is used to differentiate
    if("/history?period1=1462406400&period2=1620172800&interval=1mo&filter=history&frequency=1mo&includeAdjustedClose=true" in url):
        statement = False
    else:
        statement = True
    # gets random user agent so I don't get closed as quick
    ua = UserAgent()
    header = {'User-Agent': ua.random}
    userAgent = urllib.request.Request(url, data=None, headers=header)
    # sends the info to be opened
    website = websiteAttemptor(userAgent)

    # get's the html info
    if(isinstance(website,bool) == False):
        if(statement == True):
            soup = html(website,'html.parser')
        else:
            soup = html(website,'lxml')
        return soup
    else:
        pass

# finds the info on the yahoo website
# finds the 5 year monthly adjacent close values
def ADJParse(soupADJ):
    adj_close = soupADJ.find_all("td", {"Py(10px) Pstart(10px)"})
    count = 1
    ADJList = []
    # for loop to go through all the adj for a stock
    for i in range(len(adj_close)):
        # the table iterates through 6 statistics. adj is the 5 
        if(count % 6 == 0):
            ADJDirty = soupADJ.find_all("td", {"Py(10px) Pstart(10px)"})[i - 1]
            ADJRaw = ADJDirty.span.text.replace(",","")
            ADJRaw = float(ADJRaw)
            ADJList.append(ADJRaw)
        count += 1
    print(len(ADJList), "months of data")
    return ADJList

# looks at their cashflow statements
def cashParse(soupCash):
    cashRaw = soupCash.find_all("div", {"class":"D(tbr) fi-row Bgc($hoverBgColor):h"})
    for i in range(len(cashRaw)):
        cashSpan = cashRaw[i].span.text
        infoLine = cashRaw[i].get_text(separator="/")
        cleancash = cleanSpan(cashSpan)
        # looks for specific statement info
        if("freecashflow" in cleancash):
            freeList1 = [x.replace(",", "") for x in infoLine.split("/") if "0" in x or "," in x]
            # checks if there's null values and if so just changes it to 0. Not the best method but currently I don't
            # know what else to use
            freeList = removeNull(freeList1)
        elif("issuanceofdebt" in cleancash):
            issuanceDebt1 = [x.replace(",", "") for x in infoLine.split("/") if "0" in x or "," in x or "-" in x]
            issuanceDebt = removeNull(issuanceDebt1)
        elif("endcashposition" in cleancash):
            endCashPosition1 = [x.replace(",", "") for x in infoLine.split("/") if "0" in x or "," in x]
            endCashPosition = removeNull(endCashPosition1)
    # finds the dates that these statements were release
    dateRaw = soupCash.find_all("div", {"class":"D(tbr) C($primaryColor)"})[0]
    datecash = dateClean(dateRaw.text)
    return freeList, issuanceDebt, endCashPosition, datecash

# same thing as the free cash flow but uses the income statement instead
def incomeParse(soupIncome):
    incomeRaw = soupIncome.find_all("div", {"class":"D(tbr) fi-row Bgc($hoverBgColor):h"})
    for i in range(len(incomeRaw)):
        incomeSpan = incomeRaw[i].span.text
        infoLine = incomeRaw[i].get_text(separator="/")
        cleanIncome = cleanSpan(incomeSpan)
        if("totalrevenue" in cleanIncome):
            revenueList1 = [x.replace(",", "") for x in infoLine.split("/") if "0" in x or "," in x]
            revenueList = removeNull(revenueList1)
        elif(cleanIncome in "ebit"):
            EBIT1 = [x.replace(",", "") for x in infoLine.split("/") if "0" in x or "," in x]
            EBIT = removeNull(EBIT1)
    try:
        return revenueList, EBIT
    except:
        # some companies don't use EBIT so I just replace it if it's not available since it's not completely
        # neccessary
        EBIT = revenueList
        return revenueList, EBIT
        

# this is the function that calls all the math functions.
def cashMath(freeList, issuanceDebt, endCashPosition, dateList, revenueList, EBIT, ADJList):
    yearReturnList, weightReg, currentReturn = ADJYearMath(ADJList,dateList)
    freeList, issuanceDebt, endCashPosition, revenueList, EBIT = errorFix(freeList, issuanceDebt, endCashPosition, revenueList, EBIT, yearReturnList)
    weight = weightSolver(freeList, issuanceDebt, endCashPosition, dateList, revenueList, EBIT,yearReturnList)
    weightedList = weightedCreator(weight,freeList, issuanceDebt, endCashPosition, dateList, revenueList, EBIT)
    predictedMonth, y_predict, currentReturn = pearson(yearReturnList,weightedList, currentReturn)
    return predictedMonth, y_predict, currentReturn


# since the solver equation requires that all statements have to be of equal length. This makes sure that they are
def errorFix(freeList, issuanceDebt, endCashPosition, revenueList, EBIT, ADJList):
    totalMonth = len(ADJList) + 1
    if(len(freeList) > totalMonth):
        del freeList[totalMonth:]
    if(len(issuanceDebt) > totalMonth):
        del issuanceDebt[totalMonth:]
    if(len(endCashPosition) > totalMonth):
        del endCashPosition[totalMonth:]
    if(len(revenueList) > totalMonth):
        del revenueList[totalMonth:]
    if(len(EBIT) > totalMonth):
        del EBIT[totalMonth:]
    if(len(freeList) < len(revenueList)):
        freeList = revenueList
    if(len(issuanceDebt) < len(revenueList)):
        issuanceDebt = revenueList
    if(len(endCashPosition) < len(revenueList)):
        endCashPosition = revenueList
    if(len(EBIT) < len(revenueList)):
        EBIT = revenueList
    return freeList, issuanceDebt, endCashPosition, revenueList, EBIT


# removes null and replaces it with 0
def removeNull(nullList):
    cleanList = []
    for i in nullList:
        if(len(str(i)) == 1):
            cleanList.append(0)
        else:
            number = float(i) / 100000
            cleanList.append(number)
    return cleanList


# gets the weights created by the solver and creates a new list with them
def weightedCreator(weight,freeList, issuanceDebt, endCashPosition, dateList, revenueList, EBIT):
    weightedList = []
    for i in range(len(freeList)):
        revI = revenueList[i] * weight[0]
        ebitI = EBIT[i] * weight[1]
        freeI = freeList[i] * weight[2]
        debtI = issuanceDebt[i] * weight[3]
        endCashI = endCashPosition[i] * weight[4]
        weightedAvg = revI + ebitI + freeI + debtI + endCashI
        weightedList.append(weightedAvg)
    return weightedList



# calls the optimization and creates the dictionary used to create the df
def weightSolver(freeList, issuanceDebt, endCashPosition, dateList, revenueList, EBIT, yearReturnList):
    solverDic = {"rev": revenueList[1:], "ebit": EBIT[1:], "free": freeList[1:], "debt": issuanceDebt[1:], "endcash": endCashPosition[1:]}
    rows = len(yearReturnList)
    names = ['rev', 'ebit', "free", "debt", "endcash"]
    df_returns = dataFrameCreator(rows,names, solverDic)
    x0 = np.array([0.2,0.2,0.2,0.2,0.2])
    # I tried out all the scipy optimization methods and cobyla seemed to be the best.
    # no need of adding constraints
    out = minimize(pearsonSolver, x0, method='COBYLA', args=(df_returns,yearReturnList,))
    return out.x

# creates the df
def dataFrameCreator(rows, names, d):
    listVars= names
    rng = pd.date_range("2021", periods=rows, freq='Y')
    df_temp = pd.DataFrame(d, columns=listVars)
    df_temp = df_temp.set_index(rng)
    print(df_temp)
    return df_temp

# does the solver to find the largest r value
# I do this to get the best estimate on how a company was valued before
def pearsonSolver(weight,df, yearReturnList):
    yearReturnList = array(yearReturnList)
    weights = [weight[0],weight[1],weight[2],weight[3],weight[4]]
    r = 0
    lists = []
    for i in range(len(yearReturnList)):
        lists.append(np.dot(np.asarray(weights).T, df.iloc[i]))
    lists = array(lists)
    r = np.corrcoef(lists,yearReturnList)[0,1]
    r = abs(r)
    return -r

# with the optimal list it completes the linear regression model
def pearson(yearReturnList,revenueList, currentReturn):
    yearReturnList = array(yearReturnList)
    revenueList = array(revenueList[1:])
    model = LinearRegression()
    yearReturnList = yearReturnList.reshape(-1,1)
    revenueList = revenueList.reshape(-1,1)
    model.fit(revenueList,yearReturnList)
    X_predict = array(revenueList[0])
    X_predict = X_predict.reshape(1,-1)
    y_predict = model.predict(X_predict)
    compare = ((y_predict / currentReturn) - 1) * 100
    print(y_predict,"---------------", compare, "% return-------------------------", currentReturn)
    return compare, y_predict, currentReturn


# uses the current data and the latest financial statement report to find how long it's been since the latest financial report
# This is required so it uses the accurate adj price for each statement
def ADJYearMath(ADJList,dateList):
    now = datetime.datetime.now()
    yearNow = now.year
    monthNow = now.month
    month, day, year = dateList[0].split("/")
    monthLong = (12 - int(month)) + monthNow
    totalLong = monthLong
    currentReturn = ADJList[0]
    otherRange = len(dateList)
    weightReg = 12 / totalLong
    yearReturnList = []
    count = totalLong
    # gets the ADJ price when each statement was released
    for i in range(int(otherRange)):
        if(count < len(ADJList) - totalLong):
            yearReturn = ADJList[count]
            yearReturnList.append(yearReturn)
        count += 12
    return yearReturnList, weightReg, currentReturn


# cleans the date list
def dateClean(dateRaw):
    date = ""
    dateList = []
    count = 0
    for i in range(len(dateRaw)):
        if(count < len(dateRaw) - 1):
            if(dateRaw[count].isdigit()):
                date = date + str(dateRaw[count])
            elif(dateRaw[count] == "/"):
                date = date + str(dateRaw[count])
                if(count < len(dateRaw) - 4 and dateRaw[count + 3] != "/"):
                    date = date + str(dateRaw[count + 1]) + str(dateRaw[count + 2]) + str(dateRaw[count + 3]) + str(dateRaw[count + 4])
                    dateList.append(date)
                    date = ""
                    count += 4
            else:
                date = date
            count += 1
    return dateList



# replaces the commas in the statement balances so it can convert to float
def cashList(infoLine):
    infoLine = str(infoLine)
    cashDigit = ""
    count = 0
    cashList = [float(x.replace(",", "")) for x in infoLine.split("/") if "," in x]
    return cashList

# cleans the span so it can know what statement is the current statement
def cleanSpan(cash):
    cleancash = ""
    for i in cash:
        if(i == " "):
            cleancash = cleancash
        else:
            cleancash = cleancash + i.lower()
    return cleancash

# searches the website. Uses a proxy 
def websiteAttemptor(website):
    proxytest = proxy()
    authinfo = urllib.request.HTTPBasicAuthHandler()

    proxy_support = urllib.request.ProxyHandler({"http" : proxy()})

    # build a new opener that adds authentication and caching FTP handlers
    opener = urllib.request.build_opener(proxy_support, authinfo,
                                        urllib.request.CacheFTPHandler)

    # install it
    urllib.request.install_opener(opener)
    url = urllib.request.urlopen(website)
    return url

# creates a random proxy so the yahoo has a harder time telling where the requests are coming from
def proxy():
    randomProxy = random.randint(0,10)
    randomLetter = "abcdefghijklmnopqrstuvwxyz"
    randomNum = "0123456789"
    word = ""
    for _ in range(randomProxy):
        randomLetterNum = random.randint(0,25)
        word = randomLetter[randomLetterNum] + word
    number = ""
    for _ in range(4):
        randomNumber = random.randint(0,9)
        number = str(randomNum[randomNumber]) + str(number)
    return "http://" + word + ":" + number

