Import-Module Selenium

# Excel Workbook Path
$File = '\\ecsg\Data_Area\Site_South_Shore\SSH_2023_Inventory_AllAssets.xlsx'
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $true
$Workbook = $Excel.workbooks.open($File)
$Worksheet = $Workbook.Worksheets.Item(1)

$Range = $Worksheet.Range("A1").EntireColumn

$Driver = Start-SeEdge -Maximized
Enter-SeUrl https://navigator.americannational.com/ -Driver $Driver

Clear-Host

Write-Host "Starting John's automation script!"
$SearchItem = ""

While ($SearchItem -ne "exit") {
	$SearchItem = Read-Host -Prompt "Enter an Asset Tag (e.g.: CTS31074)"
	$Search = $Range.find($SearchItem)
	
	If ($Range.find($SearchItem) -ne $Null) {
		Write-Host "Found $SearchItem in the spreadsheet."
	 	# $Worksheet.Cells($Search.Row, 4).Value() = "Today's Date"
		$Manufacturer = $Worksheet.Cells($Search.Row, 7).Value()
		$SerialNumber = $Worksheet.Cells($Search.Row, 9).Value()
		$Custodian = $Worksheet.Cells($Search.Row, 10).Value()
		$HardwareAssetStatus = $Worksheet.Cells($Search.Row, 11).Value()
		$HardwareAssetType = $Worksheet.Cells($Search.Row, 12).Value()
		$State = $Worksheet.Cells($Search.Row, 13).Value()
		$City = $Worksheet.Cells($Search.Row, 14).Value()
		$Building = $Worksheet.Cells($Search.Row, 15).Value()
		$Floor = $Worksheet.Cells($Search.Row, 16).Value()
		$Office = $Worksheet.Cells($Search.Row, 17).Value()
	}
	ElseIf ($SearchItem -eq "exit"){
		Write-Host "Exit called -- ending the script."	
  	} 
	Else {
		Write-Host "Could not find $SearchItem in the spreadsheet."
		$excel.Quit()
		Exit
	}
	
	$LocationTag = ""
	If ($Floor -ne $Null) {
		$LocationTag = "$State-$City-$Building-$Floor"
	}
	Else {
		$LocationTag = "$State-$City-$Building"
	}
	
	$Element = Find-SeElement -Driver $Driver -Id "global-search__header__form__input_id"
	Send-SeKeys -Element $Element -Keys "$SearchItem"
	
	# $myshell = New-Object -com "Wscript.Shell"
	# $myshell.sendkeys("{ENTER}")
	
	$Element = Find-SeElement -Driver $Driver -Id "checkAllClassFilter"
	Invoke-SeClick -Element $Element
	
	Start-Sleep -Seconds 3
	
	$Element = Find-SeElement -Driver $Driver -ClassName "results-details-highlight"
	Invoke-SeClick -Element $Element
	
	Start-Sleep -Seconds 2
	
	# Custodian
	$Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasPrimaryUser"
	$Element = $Element.Clear()
	If ($Custodian -ne $Null) {
		$Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasPrimaryUser"
		Send-SeKeys -Element $Element -Keys $Custodian
	}
	
	# Location
	$Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasLocation" 
	$Element = $Element.Clear()
	$Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasLocation" 
	Send-SeKeys -Element $Element -Keys $LocationTag
	
	# Office Location Details
	$Element = Find-SeElement -Driver $Driver -Name "LocationDetails"
	$Element = $Element.Clear()
	If ($Office -ne $Null) {
		$Element = Find-SeElement -Driver $Driver -Name "LocationDetails"
		Send-SeKeys -Element $Element -Keys $Office
	}
	
	Write-Host "Script complete!"
	Write-Host `n`n`n`n
	Write-Host $SearchItem $SerialNumber $Manufacturer $SerialNumber $Custodian $HardwareAssetStatus $HardwareAssetType $State $City $Building $Floor $Office
	Write-Host `n`n`n`n
}
Write-Host "Ending the script."
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable Excel
