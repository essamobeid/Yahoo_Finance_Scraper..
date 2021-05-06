from selenium import webdriver
import time 
from bs4 import BeautifulSoup
import pandas as pd
import xlsxwriter
import matplotlib.pyplot as plt

#where PhantomJS driver is installed on your pc
PhantomJS_PATH = r"C:\webDriver\phantomjs.exe"

#This class to draw two graphs based on data from the financial statements 
class Graphs:
    
    #constructor
    def __init__(self, ticker):
        self.ticker = ticker

    #PhantomJS_PATH = r"C:\webDriver\hantomjs.exe"
    
    #this function to draw graph based on informarion from the income statement 
    def income_statement_graph(self, ticker):
        #starting the driver
        driver =webdriver.PhantomJS(PhantomJS_PATH)
        #navigating to the income statement
        driver.get("https://ca.finance.yahoo.com/quote/"+ticker+"/financials?p="+ticker+"&.tsrc=fin-srch")

        #saving the webpage data in HTML format
        html = driver.page_source
        #defining BeautifulSoup object to parse the html data
        soup = BeautifulSoup(html, "html.parser")
        ls =[]
        #looking for all divs in the html 
        for l in soup.find_all("div"):
            ls.append(l.string)
        
        #filtering the parsed data
        ls = [e for e in ls if e not in ("Operating Expenses","Non-recurring Events")]
        new_ls = list(filter(None,ls))
        temp = []
        #converting unicode to string format
        for x in new_ls:
            temp.append(x.encode('ascii', 'ignore'))
        stats =[]
        #start the income table from Total Revenue 
        stats = temp[temp.index('Total Revenue'):]
        #grouping each 6 element into a list and creating a list of lists
        #each list has the tag and the data over the past four years
        final_stats_list = [stats[i:i + 6] for i in xrange(0, (len(stats) -3), 6)] 
        
        #the ylabel 
        dates = ("2016", "2017", "2018", "2019")

        #Finding Net Income, Gross Profit, Total Revenue, and Total Operating Expenses for the graph 
        for i in range(0, (len(final_stats_list)-1)):
            if final_stats_list[i][0]== 'Net Income':
                net_income = final_stats_list[i]
            if final_stats_list[i][0] == 'Gross Profit':
                gross_profit = final_stats_list[i]
            if final_stats_list[i][0]=='Total Revenue':
                total_revenue = final_stats_list[i]
            if final_stats_list[i][0]=='Total Operating Expenses':
                operating_expenses = final_stats_list[i]
        
        #Parsing the data
        net_income=net_income[2:]
        net_income.reverse()
        gross_profit = gross_profit[2:]
        gross_profit.reverse()
        total_revenue = total_revenue[2:]
        total_revenue.reverse()
        operating_expenses = operating_expenses[2:]
        operating_expenses.reverse()

        gross_profit_int =[]
        net_income_int = []
        total_revenue_int = []
        operating_expenses_int=[]
        
        #converting strings to integers 
        for i in net_income:
            net_income_int.append(i.replace(',',''))
        for i in gross_profit:
            gross_profit_int.append(i.replace(',',''))
        for i in total_revenue:
            total_revenue_int.append(i.replace(',',''))
        for i in operating_expenses:
            operating_expenses_int.append(i.replace(',',''))
        
        gross_profit_int = map(int, gross_profit_int)
        net_income_int = map(int, net_income_int)
        total_revenue_int = map(int, total_revenue_int)
        operating_expenses_int = map(int, operating_expenses_int)

        #drawing the graphs 
        plt.plot( dates, gross_profit_int, label='Gross Profit')
        plt.plot( dates, net_income_int, label='Net Income')
        plt.plot(dates, total_revenue_int, label='Total Revenue')
        plt.plot(dates, operating_expenses_int, label="Total Operating Expenses")

        plt.xlabel('Years')
        plt.title('Income Statement')
        plt.legend(loc='best')
        plt.savefig('{}_income_statement.png'.format(ticker))

    #This function for drawing a graph based on data from the balance sheet
    def balance_sheet_graph(self, ticker):
        driver =webdriver.PhantomJS(PhantomJS_PATH)
        driver.get("https://ca.finance.yahoo.com/quote/"+ticker+"/balance-sheet?p="+ticker+"&.tsrc=fin-srch")

        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        ls =[]
        for l in soup.find_all("div"):
            ls.append(l.string)

        ls = [e for e in ls if e not in ("Operating Expenses","Non-recurring Events")]
        new_ls = list(filter(None,ls))
        temp = []
        for x in new_ls:
            temp.append(x.encode('ascii', 'ignore'))
        stats =[]
        stats = temp[temp.index('Cash And Cash Equivalents'):]
        final_stats_list = [stats[i:i + 5] for i in xrange(0, (len(stats) -3), 5)] 

        dates = ("2016", "2017", "2018", "2019")
        for i in range(0, (len(final_stats_list)-1)):
            if final_stats_list[i][0]== 'Total Current Assets':
                current_assets = final_stats_list[i]
            if final_stats_list[i][0] == 'Total Current Liabilities':
                current_liabilities = final_stats_list[i]
            if final_stats_list[i][0]=='Total Assets':
                total_assets = final_stats_list[i]
            if final_stats_list[i][0]=='Total Liabilities':
                total_liabilities = final_stats_list[i]
            if final_stats_list[i][0]=="Total stockholders' equity":
                stockholder_equity = final_stats_list[i]
        
        current_assets.remove(current_assets[0])
        current_assets.reverse()
         
        #current_liabilities = final_stats_list[21]
        current_liabilities.remove(current_liabilities[0])
        current_liabilities.reverse()

        #total_assets = final_stats_list[15]
        total_assets.remove(total_assets[0])
        total_assets.reverse()

        #total_liabilities = final_stats_list[27]
        total_liabilities.remove(total_liabilities[0])
        total_liabilities.reverse()

        #stockholder_equity = final_stats_list[31]
        stockholder_equity.remove(stockholder_equity[0])
        stockholder_equity.reverse()

        current_assets_int = []
        current_liabilities_int = []
        total_assets_int = []
        total_liabilities_int =[]
        stockholder_equity_int = []

        for i in current_assets:
            current_assets_int.append(i.replace(',',''))
        for i in current_liabilities:
            current_liabilities_int.append(i.replace(',',''))
        for i in total_assets:
            total_assets_int.append(i.replace(',',''))
        for i in total_liabilities:
            total_liabilities_int.append(i.replace(',',''))
        for i in stockholder_equity:
            stockholder_equity_int.append(i.replace(',',''))

        current_assets_int = map(int, current_assets_int)
        current_liabilities_int = map(int, current_liabilities_int)
        total_assets_int = map(int,total_assets_int)
        total_liabilities_int = map(int,total_liabilities_int)
        stockholder_equity_int = map(int,stockholder_equity_int)

        plt.plot( dates, current_assets_int, label='Total Current Assets')
        plt.plot( dates, current_liabilities_int, label='Total Current Liabilities')
        plt.plot( dates, total_assets_int, label='Total Assets')
        plt.plot( dates, total_liabilities_int, label='Total Liabilities')
        plt.plot( dates, stockholder_equity_int, label="Stockholder's Equity")

        plt.xlabel('Years')
        plt.title('Balance Sheet')
        plt.legend(loc='best')

        plt.savefig('{}_balance_sheet.png'.format(ticker))
        #driver.quit()
     
#This class to provide data from the summary page table and to scrape the latest news and articles about the desired stock      
class Summary:

    def __init__(self, ticker):
        self.ticker = ticker

    #PhantomJS_PATH = r"C:\webDriver\PhantomJSdriver.exe"

    def summary_page(self, ticker):
        driver =webdriver.PhantomJS(PhantomJS_PATH)
        driver.get("https://ca.finance.yahoo.com/quote/"+ticker+"?p="+ticker+"&.tsrc=fin-srch")

        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        
        summary = []
        summary_headers = ("Today Price", "Open", "Previous Close", "P/E Ratio", "EPS(TTM)")
        
        open_price = driver.find_element_by_xpath("""//*[@id="quote-summary"]/div[1]/table/tbody/tr[2]/td[2]/span""").text
        pe_ration = driver.find_element_by_xpath("""//*[@id="quote-summary"]/div[2]/table/tbody/tr[3]/td[2]/span""").text
        current_price = driver.find_element_by_xpath("""//*[@id="quote-header-info"]/div[3]/div/div/span[1]""").text
        eps= driver.find_element_by_xpath("""//*[@id="quote-summary"]/div[2]/table/tbody/tr[4]/td[2]/span""").text
        previous_close = driver.find_element_by_xpath("""//*[@id="quote-summary"]/div[1]/table/tbody/tr[1]/td[2]/span""").text
        summary.append(current_price)
        summary.append(open_price)
        summary.append(previous_close)
        summary.append(pe_ration)
        summary.append(eps)
        summary_temp =[]
        for x in summary:
                summary_temp.append(x.encode('ascii', 'ignore'))

        dict = {'Summary': summary_headers, 'Table': summary_temp}        
        summary_review = pd.DataFrame(dict)
        summary_review = summary_review[['Summary', 'Table']]
        return summary_review
        
    #This function to scrape links to the latest news related to the desired stock
    def articles(self, ticker):
        driver =webdriver.PhantomJS(PhantomJS_PATH)
        driver.get("https://ca.finance.yahoo.com/quote/"+ticker+"?p="+ticker)

        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        #Initiating the links and the titles lists 
        links= []
        temp=[]
        titles=[]

        #finding HTML class that has all the titles
        for div in soup.find(id="quoteNewsStream-0-Stream").find_all(class_="Mb(5px)"):
             titles.append(div.text)

        #finding HTML class that has all the links  
        for div in soup.find(id="quoteNewsStream-0-Stream").find_all(class_="Fw(b) Fz(18px) Lh(23px) LineClamp(2,46px) Fz(17px)--sm1024 Lh(19px)--sm1024 LineClamp(2,38px)--sm1024 mega-item-header-link Td(n) C(#0078ff):h C(#000) LineClamp(2,46px) LineClamp(2,38px)--sm1024 not-isInStreamVideoEnabled", href=True):
            temp.append((div['href']).encode('ascii', 'ignore'))
            
        for str in temp:
            if (str.startswith('/video') or str.startswith('/news') or str.startswith('/m/')):
                links.append('https://ca.finance.yahoo.com' + str)
            else:
                links.append(str)
                
        dict ={'Titles': titles, 'Links':links}
        #creating Pandas dataframe of links and articles 
        articles_table = pd.DataFrame(dict)
        articles_table = articles_table[['Titles', 'Links']]
        #driver.quit()

        return articles_table

#This class to scrape data from the financial statements and save it in Pandas dataframes      
class Financials:
    def __init__(self, ticker):
        self.ticker = ticker

    #PhantomJS_PATH = r"C:\webDriver\PhantomJSdriver.exe"
    
    #Scraping income statement
    def income_data(self, ticker):
        driver =webdriver.PhantomJS(PhantomJS_PATH)

        driver.get("https://ca.finance.yahoo.com/quote/" + ticker + "/financials?p=" + ticker)

        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        ls =[]
        for l in soup.find_all("div"):
            ls.append(l.string)

        ls = [e for e in ls if e not in ("Operating Expenses","Non-recurring Events")]
        new_ls = list(filter(None,ls))
        temp = []
        for x in new_ls:
            temp.append(x.encode('ascii', 'ignore'))
        stats =[]
        stats = temp[temp.index('Total Revenue'):]
        final_data_list = [stats[i:i + 6] for i in xrange(0, (len(stats) -3), 6)]        
        income_table = pd.DataFrame(final_data_list)
        #driver.quit()
        return income_table

    #Scraping balance sheet
    def balance_sheet_data(self, ticker):
        driver =webdriver.PhantomJS(PhantomJS_PATH)

        driver.get("https://ca.finance.yahoo.com/quote/" + ticker + "/balance-sheet?p=" + ticker)

        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        ls =[]
        for l in soup.find_all("div"):
            ls.append(l.string)

        ls = [e for e in ls if e not in ("Operating Expenses","Non-recurring Events")]
        new_ls = list(filter(None,ls))
        temp = []
        for x in new_ls:
            temp.append(x.encode('ascii', 'ignore'))
        stats =[]
        stats = temp[temp.index('Cash And Cash Equivalents'):]
        final_data_list = [stats[i:i + 5] for i in xrange(0, (len(stats) -3), 5)]        
        balance_table = pd.DataFrame(final_data_list)
        #driver.quit()
        return balance_table

    #Scraping cash flow statement 
    def cash_flow(self, ticker):
        driver =webdriver.PhantomJS(PhantomJS_PATH)

        driver.get("https://ca.finance.yahoo.com/quote/" + ticker + "/cash-flow?p=" + ticker)

        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        ls =[]
        for l in soup.find_all("div"):
            ls.append(l.string)

        ls = [e for e in ls if e not in ("Operating Expenses","Non-recurring Events")]
        new_ls = list(filter(None,ls))
        temp = []
        for x in new_ls:
            temp.append(x.encode('ascii', 'ignore'))
        stats =[]
        stats = temp[temp.index('Net Income'):]
        final_data_list = [stats[i:i + 6] for i in xrange(0, (len(stats) -3), 6)]        
        cash_flow_table = pd.DataFrame(final_data_list)
        #driver.quit()
        return cash_flow_table

#Facade class 
class Broker:

    def __init__(self, ticker):
        self.graphs = Graphs(ticker)
        self.summary = Summary(ticker)
        self.financials = Financials(ticker)

    def yahoo_finance(self, ticker):

        income_graph = self.graphs.income_statement_graph(ticker)
        balance_sheet_graph = self.graphs.balance_sheet_graph(ticker)
        summary_page= self.summary.summary_page(ticker)
        #summary_articles = self.summary.articles(ticker)
        income = self.financials.income_data(ticker)
        balance = self.financials.balance_sheet_data(ticker)
        cash = self.financials.cash_flow(ticker)

        #creating an Excel file
        writer = pd.ExcelWriter('{}_Analysis.xlsx'.format(ticker), engine='xlsxwriter')

        #Creating sheets in the Excel file
        cash.to_excel(writer, sheet_name='Cash Flow',index=False, header=False)
        balance.to_excel(writer, sheet_name="Balance Sheet", index=False, header=False)
        income.to_excel(writer, sheet_name="Income Statement", index=False, header=False)
        summary_page.to_excel(writer, sheet_name="Summary",index=False, header=False, startrow = 0, startcol = 0)
        #summary_articles.to_excel(writer, sheet_name="Summary", index = False, header=False, startrow = 0, startcol = 3)

        workbook  = writer.book
        #cells formating 
        cell_format = workbook.add_format({'bold': True, 'font_size': 12, 'border':1})
        
        #income statement Excel sheet
        income_worksheet = writer.sheets['Income Statement']
        income_worksheet.set_column("A:A",33, cell_format)
        income_worksheet.set_column("B:F", 20)
        income_worksheet.insert_image('H4', 'income_statement.png')

        #balance sheet Excel sheet
        balance_worksheet = writer.sheets['Balance Sheet']
        balance_worksheet.set_column("A:A", 33, cell_format)
        balance_worksheet.set_column("B:E",20)
        balance_worksheet.insert_image('G4', 'balance_sheet.png')
        
        #Cash flow Excel sheet
        cash_worksheet = writer.sheets['Cash Flow']
        cash_worksheet.set_column("A:A",33, cell_format)
        cash_worksheet.set_column("B:F", 20)

        #Summary page Excel sheet
        summary_worksheet = writer.sheets['Summary']
        summary_worksheet.set_column("A:A",20, cell_format)
        summary_worksheet.set_column("B:B",20)
        summary_worksheet.set_column("D:E", 80)
        summary_worksheet.set_column("D:D", 80, cell_format)

        writer.save()
        
ticker = raw_input("enter the ticker: ")
broker = Broker(ticker)
broker.yahoo_finance(ticker)
    
print("\nData has been extracted and stored in for the following stock:")
print(ticker)
print("Process completed successfully")