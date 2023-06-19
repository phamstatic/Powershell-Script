Write-Host "Starting John's automation script!"
# $SearchItem = Read-Host -Prompt "Enter an Asset Tag or Serial Number"
$SearchItem = "CTS1337"

$Driver = Start-SeEdge
Enter-SeUrl https://navigator.americannational.com/ -Driver $Driver


$Element = Find-SeElement -Driver $Driver -Id "global-search__header__form__input_id"
Send-SeKeys -Element $Element -Keys "$SearchItem"

# $myshell = New-Object -com "Wscript.Shell"
# $myshell.sendkeys("{ENTER}")

$Element = Find-SeElement -Driver $Driver -Id "checkAllClassFilter"
Invoke-SeClick -Element $Element
