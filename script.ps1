Import-Module Selenium

<# This variable holds the path to the Excel spreadsheet. #>
$excelPath = '\\ecsg\Data_Area\Site_South_Shore\A-F_Inventory.xlsx'

<# These variables hold the column numbers to their corresponding identifiers. #>
$navUpdateColumn = 4
$manufacturerColumn = 7
$serialNumberColumn = 9
$custodianColumn = 11
$hardwareAssetStatusColumn = 13
$hardwareAssetTypeColumn = 14
$stateColumn = 16
$cityColumn = 18
$buildingColumn = 19
$floorColumn = 21
$officeColumn = 22

<# This block holds the process of opening up Microsoft Edge. #>
Try {
    $Driver = Start-SeEdge -Maximized -ErrorAction Stop
    Enter-SeUrl https://navigator.americannational.com/ -Driver $Driver -ErrorAction Stop
}
Catch [System.Management.Automation.MethodInvocationException]{
    Throw "Could not find or open the Microsoft Edge browser."
    break
}
Finally {
    Clear-Host
}

<# This block holds the process of opening the Excel spreadsheet. #>
Try {
    $Excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $Excel.visible = $true
    $Workbook = $Excel.workbooks.open($excelPath)
    $Worksheet = $Workbook.Worksheets.Item(1)
}
Catch {
    Throw "Could not find or open the Microsoft Excel spreadsheet."
    break
}

$Range = $Worksheet.Range("A1").EntireColumn
$SearchItem = ""

<# This block will run the script in an infinite loop until "exit" has been inputted. The program will search for the search item in the spreadsheet and obtains all related information. #>
Write-Host "REMINDER: Double check that the information matches from the spreadsheet to Navigator before saving." -ForegroundColor yellow
While ($SearchItem -ne "exit") {
    $SearchItem = Read-Host -Prompt "Enter an Asset Tag"
    $Search = $Range.find($SearchItem, [Type]::Missing, [Type]::Missing, 1)
    If ($Null -ne $Range.find($SearchItem, [Type]::Missing, [Type]::Missing, 1)) {
        Write-Host "Found $SearchItem in the spreadsheet."
        # $Worksheet.Cells($Search.Row, $navUpdateColumn).Value() = (Get-Date).ToString('M/dd/yy') # This line will update the Nav Update column if uncommented.
		$Manufacturer = $Worksheet.Cells($Search.Row, $manufacturerColumn).Value() 
		$SerialNumber = $Worksheet.Cells($Search.Row, $serialNumberColumn).Value()
		$Custodian = $Worksheet.Cells($Search.Row, $custodianColumn).Value() 
		$HardwareAssetStatus = $Worksheet.Cells($Search.Row, $hardwareAssetStatusColumn).Value()
		$HardwareAssetType = $Worksheet.Cells($Search.Row, $hardwareAssetTypeColumn).Value()
		$State = $Worksheet.Cells($Search.Row, $stateColumn).Value()
		$City = $Worksheet.Cells($Search.Row, $cityColumn).Value()
		$Building = $Worksheet.Cells($Search.Row, $buildingColumn).Value()
		$Floor = $Worksheet.Cells($Search.Row, $floorColumn).Value()
		$Office = $Worksheet.Cells($Search.Row, $officeColumn).Value()
        Write-Host $SearchItem $SerialNumber $Manufacturer $SerialNumber $Custodian $HardwareAssetStatus $HardwareAssetType $State $City $Building $Floor $Office -ForegroundColor green -BackgroundColor black
    
        $LocationTag = ""
        If (($Null -ne $State) -and ($Null -ne $City) -and ($Null -ne $Building) -and ($Null -ne $Floor)) {
            $LocationTag = "$State-$City-$Building-$Floor"
        }
        ElseIf (($Null -ne $State) -and ($Null -ne $City) -and ($Null -ne $Building)) {
            $LocationTag = "$State-$City-$Building"
        }
        ElseIf (($Null -ne $State) -and ($Null -ne $City)) {
            $LocationTag = "$State-$City"
        }
        ElseIf (($Null -ne $State)) {
            $LocationTag = "$State"
        }

        <# This block checks and waits for the search bar on the webpage. #>
        do {
            $Failed = $false
            Try {
                $Element = Find-SeElement -Driver $Driver -Id "global-search__header__form__input_id"
                Send-SeKeys -Element $Element -Keys "$SearchItem"
            }
            Catch [System.Management.Automation.MethodInvocationException] {
                Write-Host "Failed to find the search bar... trying again (double check that the browser is not minimized)." -ForegroundColor red
                $Failed = $true
            }  
        } while ($Failed)

        <# This block checks and waits for the class filter button on the webpage. #>
        do {
            $Failed = $false
            Try {
                Start-Sleep -Seconds 1
                $Element = Find-SeElement -Driver $Driver -Id "checkAllClassFilter"
                Invoke-SeClick -Element $Element
            }
            Catch [System.Management.Automation.MethodInvocationException] {
                Write-Host "Could not find the Class filter... trying again." -ForegroundColor red
                $Failed = $true
            }
        } while ($Failed)

        <# This block checks and waits for the hardware asset link on the webpage. #>
        do {
            $Failed = $false
            Try {
                Start-Sleep -Seconds 1
                $Element = Find-SeElement -Driver $Driver -ClassName "results-details-highlight"
                Invoke-SeClick -Element $Element[$Element.length - 1]
            }
            Catch [System.Management.Automation.MethodInvocationException] {
                Write-Host "Could not find the hardware asset link... trying again." -ForegroundColor red
                $Failed = $true
            }
        } while ($Failed)

        <# This block checks and updates the Hardware Asset Status parameter. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -ClassName "k-ext-dropdown"
            $Element[4].Clear()
            Send-SeKeys -Element $Element[4] -Keys $HardwareAssetStatus
            Write-Host "Hardware Asset Status has been updated to $HardwareAssetStatus" -ForegroundColor blue
        }
        Catch {
            Write-Host "Something went wrong with inputting $HardwareAssetStatus into Hardware Asset Status." -ForegroundColor red 
        }

        <# This block checks and updates the Custodian parameter. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasPrimaryUser"
            $Element = $Element.Clear()
            If ($Null -ne $Custodian) {
                $Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasPrimaryUser"
                Send-SeKeys -Element $Element -Keys $Custodian
                Write-Host "Custodian has been updated to $Custodian" -ForegroundColor blue
            }
        }
        Catch {
            Write-Host "Something went wrong with inputting $Custodian into Custodian." -ForegroundColor red 
        }

        <# This block checks and updates the Location parameter. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasLocation" 
            $Element = $Element.Clear()
            $Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasLocation" 
            Send-SeKeys -Element $Element -Keys $LocationTag
            Write-Host "Location has been updated to $LocationTag" -ForegroundColor blue
        }
        Catch {
            Write-Host "Something went wrong with inputting $LocationTag into Location." -ForegroundColor red 
        }

        <# This block checks and updates the Location Details parameter. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -Name "LocationDetails"
            $Element = $Element.Clear()
            If ($Null -ne $Office) {
                $Element = Find-SeElement -Driver $Driver -Name "LocationDetails"
                Send-SeKeys -Element $Element -Keys $Office
                Write-Host "Location Details  has been updated to $Office" -ForegroundColor blue
            }
        }
        Catch {
            Write-Host "Something went wrong with inputting $Office into Location Details." -ForegroundColor red 
        }
    }

    ElseIf ($SearchItem -eq "exit") {
        Write-Host "Exit called -- ending the script." -ForegroundColor red
    }
    Else {
        Write-Host "Could not find $SearchItem in the spreadsheet." -ForegroundColor red
    }
}
