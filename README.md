# excelPrices
This takes price data from Auctionator.lua into a spreadsheet for only the most elite of Goblins

# Requirements
* Have and use Auctionator
* Have Powershell with installed Excel library Workbooks -> Example of the call this uses: https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open

# Usage
1. In game run an auctionator scan
0. Exit game so the data gets out of memory and persisted to the file
1. Fill a spreadsheet with a list of item names you care about getting the prices for in a single column on the first sheet (rest of sheet is empty)
2. Make other sheets that depend on the first sheet's prices will show on the second column after running script
3. Make a copy of your price sheet for the script to use
4. Make sure the first two variables in the script are pointing to the right paths (Hint, they are not because it is for my environment)
5. Run script

# Recommendations
* Make a Google Sheets document so I could edit this when I was on the go
* First sheet was the item name and prices columns
* Second was expected vendor prices (Not all friendly vendors sell at the same price - Reputation matters too)
* Third was someone else's Disenchanting table that I could do the math (inputting the current prices of mats) and determine if a crafted item was worth vendoring, disenchanting, or not doing
* Fourth+ Spreadsheet for each profession showing if there are any items to craft for a profit from the AH.
* Download a *Copy* of the sheets to use with the script for each session.

# Limitations
* I remember something having to do with the apostrophe (') not working well with my script, Do not put wow items with apostrophes in their name into the script

# Regrets
* I wish I didn't tie myself down to powershell and Auctionator, should have used Python and used Blizzard's API for getting prices. Oh well, still fun.
