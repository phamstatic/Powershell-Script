Write-Host "Starting John's automation script!"

$Driver = Start-SeEdge
Enter-SeUrl https://navigator.americannational.com/Page/f6b4a50f-6bec-4ccf-90ca-29bcd4924d30#/ -Driver $Driver

$Element = Find-SeElement -Driver $Driver -Id "global-search__header__form__input_id"
Send-SeKeys -Element $Element -Keys "SERIAL NUMBER HERE"


$Element = Find-SeElement -Driver -Id"1_gsCheckbox"
Write-Host "Finding Hardware Assets"
-Element $Element Checked "Checked"
