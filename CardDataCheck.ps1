$banner = @"
 ---------------------------------------------------
|      Card-Data Extraction Script by LatishD.      |                                                   
 ---------------------------------------------------
"@
$banner

#Fetching Computer Name
$ComName = New-Object -Typename PSCustomObject -Property @{
    ComputerName = $env:ComputerName
}
"Script running on " + $ComName

#Hardcoded path for data validation. Here you can add as many drives as you want.
$path = "D:\test\" , "D:\test1\", "C:\"

#Searching for card data in excels
$excelSheets = Get-Childitem -Path $path -Include *.xls,*.xlsx -Recurse -ErrorAction SilentlyContinue
$excel = New-Object -comobject Excel.Application
$excel.visible = $false

#Regex to identify valid cards
$masterCard = '^(?:4[0-9]{12}(?:[0-9]{3})?|5[1-5][0-9]{14})$'
$visa = '^4\d{3}(| |-)(?:\d{4}\1){2}\d{4}$'
$amex = '^3[47]\d{13,14}$'
$amexAdvanced = '^3[47]\d{1,2}(| |-)\d{6}\1\d{6}$'
$bcglobal = '^(6541|6556)[0-9]{12}$'
$carteblanche = '^389[0-9]{11}$'
$dinersClub = '^3(?:0[0-5]|[68][0-9])[0-9]{11}$'
$discover = '^65[4-9][0-9]{13}|64[4-9][0-9]{13}|6011[0-9]{12}|(622(?:12[6-9]|1[3-9][0-9]|[2-8][0-9][0-9]|9[01][0-9]|92[0-5])[0-9]{10})$'
$discoverAdvanced = '^6(?:011|5\d\d)(| |-)(?:\d{4}\1){2}\d{4}$'
$instaPayment = '^63[7-9][0-9]{13}$'
$jcb = '^(?:2131|1800|35\d{3})\d{11}$'
$koreanLocal = '^9[0-9]{15}$'
$laser = '^(6304|6706|6709|6771)[0-9]{12,15}$'
$maestro = '^(5018|5020|5038|6304|6759|6761|6763)[0-9]{8,15}$'
$masterCard1 = '^5[1-5][0-9]{14}$'
$mastercardAdvanced = '^5[1-5]\d{2}(| |-)(?:\d{4}\1){2}\d{4}$'
$solo = '^(6334|6767)[0-9]{12}|(6334|6767)[0-9]{14}|(6334|6767)[0-9]{15}$'
$switch = '^(4903|4905|4911|4936|6333|6759)[0-9]{12}|(4903|4905|4911|4936|6333|6759)[0-9]{14}|(4903|4905|4911|4936|6333|6759)[0-9]{15}|564182[0-9]{10}|564182[0-9]{12}|564182[0-9]{13}|633110[0-9]{10}|633110[0-9]{12}|633110[0-9]{13}$'
$unionPay = '^(62[0-9]{14,17})$'
$visaMaster = '^(?:4[0-9]{12}(?:[0-9]{3})?|5[1-5][0-9]{14})$'
$visaAdvanced = '^4\d{3}(| |-)(?:\d{4}\1){2}\d{4}$'

foreach($excelSheet in $excelSheets)
{
 $workbook = $excel.Workbooks.Open($excelSheet)
 "--------------------------------------------------------------"
 "[+] There are $($workbook.Sheets.count) sheets in $excelSheet"
 "--------------------------------------------------------------"

 For($i = 1 ; $i -le $workbook.Sheets.count ; $i++)
 {
  $worksheet = $workbook.sheets.item($i)
  "[i] Looking for sensitive data in $($worksheet.name) worksheet"
  $rowMax = ($worksheet.usedRange.rows).count
  $columnMax = ($worksheet.usedRange.columns).count
  For($row = 1 ; $row -le $rowMax ; $row ++)
  {
   For($column = 1 ; $column -le $columnMax ; $column ++)
    {
     [string]$formula = $workSheet.cells.item($row,$column).formula
     if(($formula -match $masterCard) -or ($formula -match $masterCard1) -or ($formula -match $mastercardAdvanced) -or ($formula -match $visaAdvanced) -or ($formula -match $unionPay) -or ($formula -match $visaMaster) -or ($formula -match $solo) -or ($formula -match $maestro) -or ($formula -match $laser) -or ($formula -match $koreanLocal) -or ($formula -match $instaPayment) -or ($formula -match $jcb) -or ($formula -match $discoverAdvanced) -or ($formula -match $discover) -or ($formula -match $carteblanche) -or ($formula -match $bcglobal) -or ($formula -match $visa) -or ($formula -match $amexAdvanced) -or ($formula -match $amex) -or ($formula -match $dinersClub)) 
     {"`t[>] Card data exposed: $($formula)"}
    } #end for $column
   } #end for $row
  $worksheet = $rowmax = $columnMax = $row = $column = $formula = $null
 } #end for
 $workbook.saved = $true
 $workbook.close()
} #end foreach
 
 $excel.quit()
 $excel = $null
 [gc]::collect()
 [gc]::WaitForPendingFinalizers()

