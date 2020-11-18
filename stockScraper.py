import os, requests, re, time
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from bs4 import BeautifulSoup
from progress.bar import IncrementalBar

# Set absolute file path for script and relative path for ticker file
script_path = os.path.dirname(__file__)
text_path = os.path.join(script_path, 'tickerList.txt')
centered = Alignment(horizontal="center", vertical="center")

# URL's for sites to be scraped
guru_foc= 'https://www.gurufocus.com/stock/'
mark_watch = 'https://www.marketwatch.com/investing/stock/'

title_font = Font(size = "14", bold = True, name = 'Calibri')
label_font = Font(name= 'Calibri', bold = True)
reg_font = Font(name = 'Calibri')

# Fill values
red = PatternFill("solid", fgColor="ffa6a6")
green = PatternFill("solid", fgColor="afd095")
yellow = PatternFill("solid", fgColor="ffffa6")
blue = PatternFill("solid", fgColor="b4c7dc")
orange = PatternFill("solid", fgColor="ffb66c")
purple = PatternFill("solid", fgColor="bf819e")

# Values for data scraping stockOutput
m_watch_out = []
g_focus_out = []

count = 1
f_count = 0
fill_list = [red, green, yellow, blue, orange, purple]

# Get list of tickers
ticker_list = []
def get_tickers(path):
    ticker_file = open(path)
    string = ticker_file.read()
    pattern = re.compile(r'[a-zA-Z]+')
    find_result = pattern.findall(string)
    for ticker in find_result:
        ticker_list.append(ticker)
    ticker_file.close()

def get_m_watch(url):
    res = requests.get(url)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, 'html.parser')
    title = soup.select('body > div.container.container--body > div.region.region--intraday > div:nth-child(2) > div > div:nth-child(2) > h1')
    price = soup.select('body > div.container.container--body > div.region.region--intraday > div.column.column--aside > div > div.intraday__data > h3 > bg-quote')
    eps = soup.select('body > div.container.container--body > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(10) > span.primary')
    beta = soup.select('body > div.container.container--body > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(7) > span.primary')
    div = soup.select('body > div.container.container--body > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(12) > span.primary')
    m_watch_out.append(title[0].text)
    m_watch_out.append(price[0].text)
    m_watch_out.append(eps[0].text)
    m_watch_out.append(beta[0].text)
    m_watch_out.append(div[0].text)
    
def get_g_focus(url):
    res = requests.get(url)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, 'html.parser')
    dte = soup.select('#financial-strength > div > table.stock-indicator-table > tbody > tr:nth-child(1) > td:nth-child(2)')
    pb = soup.select('#ratios > div > table.stock-indicator-table > tbody > tr:nth-child(3) > td:nth-child(2)')
    div_ratio = soup.select('#dividend > div > table.stock-indicator-table > tbody > tr:nth-child(1) > td:nth-child(2)')
    g_focus_out.append(dte[0].text)
    g_focus_out.append(pb[0].text)
    g_focus_out.append(div_ratio[0].text)

def scrape(a,b):
    # Creates and populates cells in the spreadsheet
    # arguments are for the y range
    get_m_watch(mark_watch + ticker)
    get_g_focus(guru_foc + ticker)
    for x in range(count, count + 9):
        for y in range(a, b):
            if x == count: # Title cell
                if y == a:
                    cell = ws.cell(row=x, column=y, value= m_watch_out[0]) 
                    cell.font = title_font
                    cell.fill = fill_list[f_count]
            if x == count + 1: # Price cell
                if y == a:
                    cell = ws.cell(row=x, column=y, value= "$" + m_watch_out[1])
                    cell.font = label_font
                    cell.fill = fill_list[f_count]
            if x == count + 2: # Debt to equity
                if y == a:
                    cell = ws.cell(row=x, column=y, value= 'Debt To Equity')
                    cell.font = reg_font
                    cell.fill = fill_list[f_count]
                if y == a + 1:
                    cell = ws.cell(row=x, column=y, value= g_focus_out[0])
                    cell.alignment = centered
                    cell.font = label_font
                    cell.fill = fill_list[f_count]
            if x == count + 3: # Earnins Per Share
                if y == a:
                    cell = ws.cell(row=x, column=y, value= 'Earnings Per Share')
                    cell.font = reg_font
                    cell.fill = fill_list[f_count]
                if y == a + 1:
                    cell = ws.cell(row=x, column=y, value= m_watch_out[2])
                    cell.alignment = centered
                    cell.font = label_font
                    cell.fill = fill_list[f_count]
            if x == count + 4: # Beta Ratio
                if y == a:
                    cell = ws.cell(row=x, column=y, value= 'Beta Ratio')
                    cell.font = reg_font
                    cell.fill = fill_list[f_count]
                if y == a + 1:
                    cell = ws.cell(row=x, column=y, value= m_watch_out[3])
                    cell.alignment = centered
                    cell.font = label_font
                    cell.fill = fill_list[f_count]
            if x == count + 5: # Price to Book
                if y == a:
                    cell = ws.cell(row=x, column=y, value= 'Price to Book')
                    cell.font = reg_font
                    cell.fill = fill_list[f_count]
                if y == a + 1:
                    cell = ws.cell(row=x, column=y, value= g_focus_out[1])
                    cell.alignment = centered
                    cell.font = label_font
                    cell.fill = fill_list[f_count]
            if x == count + 6: # Dividend Payment
                if y == a:
                    cell = ws.cell(row=x, column=y, value= 'Dividend')
                    cell.font = reg_font
                    cell.fill = fill_list[f_count]
                if y == a + 1:
                    cell = ws.cell(row=x, column=y, value= m_watch_out[4])
                    cell.alignment = centered
                    cell.font = label_font
                    cell.fill = fill_list[f_count]
            if x == count + 7: # Dividend Yield %
                if y == a:
                    cell = ws.cell(row=x, column=y, value= 'Dividend Yield %')
                    cell.font = reg_font
                    cell.fill = fill_list[f_count]
                if y == a + 1:
                    cell = ws.cell(row=x, column=y, value= g_focus_out[2])
                    cell.alignment = centered
                    cell.font = label_font
                    cell.fill = fill_list[f_count]
    m_watch_out.clear()
    g_focus_out.clear()

# Create Workbook to be saved to "stockOutput.xlsx"
wb = Workbook()
ws = wb.active # Select default sheet
ws.title = "stockOutput" # Rename sheet title 

get_tickers(text_path) # Populate ticker list 
bar = IncrementalBar('Getting Stock Info', max = len(ticker_list)) # Scraping Stock Bar

for ticker in ticker_list[:len(ticker_list) // 2]:
    scrape(1,3) # Execute scraper function
    count += 9
    if f_count < 5:
        f_count += 1
    else:
        f_count = 0
    bar.next() 
    time.sleep(1)
    
count = 1

for ticker in ticker_list[len(ticker_list) // 2:]:
    scrape(4,6) # Execute scraper function
    count += 9
    if f_count < 5:
        f_count += 1
    else:
        f_count = 0
    bar.next() 
    time.sleep(1)

bar.finish() # Show finished with progress bar

ws.column_dimensions['a'].width = 50 # Set Column A width
ws.column_dimensions['b'].width = 10 # Set Column B width
ws.column_dimensions['d'].width = 50 # Set Column D width
ws.column_dimensions['e'].width = 10 # Set Column E width

wb.save("stockOutput.xlsx") # Save out file. NOTE: Overwrites entire file everytime script runs.
