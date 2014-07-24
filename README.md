UCL Thomson Reuters Eikon Data Extraction
==


** Author : Yuanjia Su  **

** Contact: yuanjia.su.13 at ucl dot ac dot uk **


eikoncrawl- version 1.0 
---

<br/>The script is built in Excel and Eikon, with the following functionalities in order:


1. extract different indecies in 16 European financial markets 
2. specify stock currency, capitalisaiton, location, etc.
3. extract intraday data (high, open, volume, etc.)
4. extract tick data (bid, ask, trade) of each stock in any given day 
5. calculate liquidity metrics (bid/ask spread, VAMP, time difference, ect.) on tick data 
6. auto check and fill missing dates
7. auto check unretrived data and retry 3 times 
8. auto save every 5 mins  
9. logging status is visible as the program runs 


*24/07 UPDATES:*

- Refined metrics 
- Excel formulas replace VBA functions, hence more efficient  
- more market indecies are added

The script includes following three files, which are:

- main.bas
- calcmetrics.bas 
- data.bas 


Dependencies
---

- Microsoft Excei
- Thomson Reuters Eikon
- Windows Environment


Examples
---

- Extract FTSE100 index with corresponding constituent names, and save them on 
"sheet1" cell "A1"
    
```
   	Worksheet("sheet1"").Range("A1").Formula = _
    "=TR("".FTSE"",""TR.IndexConstituentRIC"",""RH=In"")" 
```


- Extract stock 'BP.L' currency, market capitalisation, and exchange country. 

```
	Worksheets("Sheet2").Range("C1:C10").Formula = _
	"=TR("BP.L",""CURRENCY;TR.CompanyMarketCap;TR.ExchangeCountry"")"
```

- Calculate market cap type 3 seconds later, ie. set enough time to guarantee data retrieving will finish. 

```
    Call getStockPrices
    startTime = Now + TimeValue("00:00:3")
    Application.OnTime startTime, "getCAP"
```


The MIT License (MIT)
---

Copyright (c) 2014 Yuanjia Su

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


