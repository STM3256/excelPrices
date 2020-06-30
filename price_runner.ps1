#This script grabs list of items in my excel, looks them up in the Auctioneer.lua file to get prices, then puts the prices into the same sheet that it got them from.
#To use this effectively, open wow, scan the AH with Auctionator, then exit wow. The LUA file will get updated after closing the application. 
#THEN run this script to update your spreadsheet
#prof0ak
#10-16-2019

# This expects your spreadsheet's first sheet to be only two columns
# First Row = [ ITEMS ] , [ PRICE ]
# The A column is the name of items you care about as input
# The B column is the output of this script
# make other sheets that depend on those prices

$excel_path = "C:\Users\prof0ak\Downloads\WOW_economy_2.xlsx"
$auctionator_path = 'D:\Games\Blizzard\World of Warcraft\_classic_\WTF\Account\prof0ak\SavedVariables\Auctionator.lua'

$items = @()
$prices = @()
foreach ($item in (Import-XLSX -Path $excel_path -RowStart 1)){
$items += $item.Items
}
Write-Host -NoNewline 'Retrieving Data for '$items.Count' items.'
Write-Host ' '
Write-Host ' --------------------'

$full = Get-Content $auctionator_path -Raw
$db_start = $full.IndexOf("AUCTIONATOR_PRICE_DATABASE = {")
$db_end = $full.IndexOf("AUCTIONATOR_STACKING_PREFS = {")
$lua = $full.Substring($db_start, $db_end-$db_start)

$end = '},'
$mr_str = '["mr"] = '
$mr_str_length = $mr_str.Length
$search_prepend = '["'
$search_append = '"] = {'

foreach ($item in $items){
#get price
$item_title = $search_prepend.ToString()+$item.ToString()+$search_append.ToString()
$start_loc = $lua.IndexOf($item_title) + $item_title.Length
$end_loc = $lua.IndexOf($end, $start_loc)
$numChar = $end_loc - $start_loc
$values = $lua.Substring($start_loc, $numChar)
$mr = $values.IndexOf($mr_str)

if($mr -lt 0){
    $price = 'null'
} else {
    $mr_end = $values.IndexOf(',', $mr)
    $mr_value_length = $mr_end - ($mr + $mr_str_length)
    $price = $values.Substring($mr + $mr_str_length, $mr_value_length)
}

Write-Host -NoNewline 'price of '$item' is: '$price
Write-Host

#add price to array
$prices += $price
}

$xl=New-Object -ComObject Excel.Application
$wb=$xl.WorkBooks.Open($excel_path)
$ws=$wb.WorkSheets.item(1)
$xl.Visible=$true

$row = 2
foreach ($price in $prices){
    $ws.Cells.Item($row,2)=$price
    $row += 1
}
$wb.Save()
$xl.Quit()
