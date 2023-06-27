Import-Module Selenium

# Excel Workbook Path
$File = '\\ecsg\Data_Area\Site_South_Shore\A-F_Inventory.xlsx'
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $true
$Workbook = $Excel.workbooks.open($File)
$Worksheet = $Workbook.Worksheets.Item(1)

$Range = $Worksheet.Range("A1").EntireColumn

$Driver = Start-SeEdge -Maximized
Enter-SeUrl https://navigator.americannational.com/ -Driver $Driver

Clear-Host

Write-Host "Starting John's automation script!"
Write-Host "Today's date is " (Get-Date).ToString('M/dd/yy')
Write-Host "The current file path is $File"

$SearchItem = ""
While ($SearchItem -ne "exit") {
	$SearchItem = Read-Host -Prompt "Enter an Asset Tag"
	$Search = $Range.find($SearchItem, [Type]::Missing, [Type]::Missing, 1)
	
	If ($Null -ne $Range.find($SearchItem, [Type]::Missing, [Type]::Missing, 1)) {
		Write-Host "Found $SearchItem in the spreadsheet."
	 	$Worksheet.Cells($Search.Row, 4).Value() = (Get-Date).ToString('M/dd/yy') # Update Nav Update Column to Today's Date
		$Manufacturer = $Worksheet.Cells($Search.Row, 7).Value() 
		$SerialNumber = $Worksheet.Cells($Search.Row, 9).Value()
		$Custodian = $Worksheet.Cells($Search.Row, 11).Value() 
		$HardwareAssetStatus = $Worksheet.Cells($Search.Row, 13).Value()
		$HardwareAssetType = $Worksheet.Cells($Search.Row, 14).Value()
		$State = $Worksheet.Cells($Search.Row, 16).Value()
		$City = $Worksheet.Cells($Search.Row, 18).Value()
		$Building = $Worksheet.Cells($Search.Row, 19).Value()
		$Floor = $Worksheet.Cells($Search.Row, 21).Value()
		$Office = $Worksheet.Cells($Search.Row, 22).Value()
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
	If ($Null -ne $Floor) {
		$LocationTag = "$State-$City-$Building-$Floor"
	}
	Else {
		$LocationTag = "$State-$City-$Building"
	}
	
	$Element = Find-SeElement -Driver $Driver -Id "global-search__header__form__input_id"
	Send-SeKeys -Element $Element -Keys "$SearchItem"
	
	$Element = Find-SeElement -Driver $Driver -Id "checkAllClassFilter"
	Invoke-SeClick -Element $Element
	Start-Sleep -Seconds 3

	$Element = Find-SeElement -Driver $Driver -ClassName "results-details-highlight"

 	Invoke-SeClick -Element $Element[$Element.length - 1]

	Start-Sleep -Seconds 2

	# Hardware Asset Status
	$Element = Find-SeElement -Driver $Driver -ClassName "k-ext-dropdown"
	$Element[4].Clear()
	Send-SeKeys -Element $Element[4] -Keys $HardwareAssetStatus

	# Custodian
	$Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasPrimaryUser"
	$Element = $Element.Clear()
	If ($Null -ne $Custodian) {
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
	If ($Null -ne $Office) {
		$Element = Find-SeElement -Driver $Driver -Name "LocationDetails"
		Send-SeKeys -Element $Element -Keys $Office
	}
	
	Write-Host "Script complete!"
	Write-Host "Double check the information on Navigator and check for Hardware Asset Status."
	Write-Host `n`
	Write-Host $SearchItem $SerialNumber $Manufacturer $SerialNumber $Custodian $HardwareAssetStatus $HardwareAssetType $State $City $Building $Floor $Office
	Write-Host `n`
}

Write-Host "Ending the script."
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable Excel