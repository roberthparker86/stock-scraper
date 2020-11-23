Takes a .txt file list of stock tickers and returns a .xlsx file with key information on each stock.

Requires Python3, as well as the modules openpyxl, BeautifulSoup 4,requests, and progress. The modules can be installed with the following commands in the terminal: 
Beautiful Soup 4: "pip install beautifulsoup4" 
Requests: "python -m pip install requests" 
Progress: "pip install progress" 
OpenPyXL: "pip install openpyxl"

To use, write out all the stock tickers one wants information on in the "tickerList.txt" file; one ticker per line. Save the file. 
NOTE: the "tickerList.txt" file must remain in the same folder as the script to be executed (stockScraper.py).

For windows, use the run command and type in the absolute path to the stock scraper files + "scrape.bat". 
For Linux, navigate via the terminal to the directory containing the script, and simply type "python3 stockScraper.py".
