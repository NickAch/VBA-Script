# VBA Stock Price Script 

![alt text](https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRNrzXT1W5HZTihHv4MfGcf0Q8lXgfNLDX3KTZUbv0BmEnnFcLXfg "Logo Title Text 1")

##Initial Commit September 15th 2014

Made by Nicholas Achkarian with some attributions to Robert from [this](http://www.financial-modelling.net/tutorials/excel/open-excel-retrieve-stock-quotes-in-a-formula) comment section for demonstrating how to Receive the value from Yahoo! Finance

##Purpose:
I have a personal "balance sheet" spreadsheet within Excel which tracks my cash, investments, cash equivalents (prepaid credit cards and gift cards),debts owed to me, credit card limits and current balances. Due to regulatory reasons I cannot link my bank accounts directly to it, but I can receive real-time updates of the value of my portfolio which contains some stocks and a mutual fund. This allows me to receive accurate representations of my assets' performance.

## How it works:
The function takes in a cell with a ticker and a cell with the total units I own and returns the total value of the security at the time. The rest of the spreadsheet gives me subtotals of cash, investments and total assets.

## To do going forward:
1. Add an option for no quantity, but rather just return the individual unit price
2. Build different flavours of the function for different stats as my portfolio gets more complex (YTD Returns, 52 week High/Low, performance comparison to similar funds)
3. Build a custom app of my own instead of tracking my portfolio through Excel