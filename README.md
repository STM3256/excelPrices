# excelPrices
This takes price data from Auctionator.lua into a spreadsheet

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

# Regrets
1. I wish I didn't tie myself down to powershell and Auctionator, should have used Python and used Blizzard's API for getting prices. Oh well, still fun.
