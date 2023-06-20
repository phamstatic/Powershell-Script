Write-Host "Starting John's automation script!"
# $SearchItem = Read-Host -Prompt "Enter an Asset Tag or Serial Number"
$SearchItem = "CTS34894"

### Need to find a way to connect to Excel and obtain the information below automatically.
$SerialNumber = "1BMYLL3"
$AssetStatus = "Deployed"
$Model = "Dell Latitude 3310 i3"
$PrimaryUser = "Pham, John"
$Location = "N/A"
$LocationDetails = "N/A"

$Driver = Start-SeEdge

Enter-SeUrl https://navigator.americannational.com/ -Driver $Driver

$Element = Find-SeElement -Driver $Driver -Id "global-search__header__form__input_id"
Send-SeKeys -Element $Element -Keys "$SearchItem"

# $myshell = New-Object -com "Wscript.Shell"
# $myshell.sendkeys("{ENTER}")

$Element = Find-SeElement -Driver $Driver -Id "checkAllClassFilter"
Invoke-SeClick -Element $Element

Write-Host "Waiting for the Hardware Asset to appear. (5 seconds)"
Start-Sleep -Seconds 5

$Element = Find-SeElement -Driver $Driver -ClassName "results-details-highlight"
Invoke-SeClick -Element $Element

