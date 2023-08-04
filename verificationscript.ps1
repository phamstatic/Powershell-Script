Import-Module $PWD\Selenium

$myInitials = "2023 Inventory JP"
$excelPath = 'Path to the Excel Spreadsheet'
$navigatorLink = 'Link to Navigator'

<# Excel Column Variables #>
$navUpdateColumn = 4
$manufacturerColumn = 7
$serialNumberColumn = 9
$custodianColumn = 11
$hardwareAssetStatusColumn = 13
$hardwareAssetTypeColumn = 14
$stateColumn = 16
$cityColumn = 18
$buildingColumn = 20
$floorColumn = 22
$officeColumn = 24

<# Open the Microsoft Edge browser. #>
Try {
    $Driver = Start-SeEdge -Maximized -ErrorAction Stop
    Enter-SeUrl $navigatorLink -Driver $Driver -ErrorAction Stop
}
Catch [System.Management.Automation.MethodInvocationException]{
    Throw "Could not find or open the Microsoft Edge browser."
    break
}
Finally {
    Clear-Host
}

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
While ($SearchItem -ne "exit") {
    Write-Host `n
    $SearchItem = Read-Host -Prompt "Enter an Asset Tag"
    $Search = $Range.find($SearchItem, [Type]::Missing, [Type]::Missing, 1)
    If ($Null -ne $Range.find($SearchItem, [Type]::Missing, [Type]::Missing, 1)) {
        Write-Host "Found $SearchItem in the spreadsheet."
        $Worksheet.Cells($Search.Row, $navUpdateColumn).Value() = (Get-Date).ToString('M/dd/yy') # This line will update the Nav Update column if uncommented.
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
        Write-Host $SearchItem $SerialNumber $Manufacturer $SerialNumber $Custodian $HardwareAssetStatus $HardwareAssetType $State $City $Building $Floor $Office -ForegroundColor yellow -BackgroundColor black
    
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

        <# This block finds the upper search bar. #>
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

        <# This block finds the check all class filter.  #>
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

        <# This block finds the Finance tab button. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -TagName a
            Invoke-SeClick -Element $Element[42]
        } 
        Catch {
            Write-Host "Failed to click the Finance tab button." -ForegroundColor red
        }

        <# This block checks and verifies the custodian within the Finance tab. #>
        do {
            $Failed = $false
            Try {
                $Element = Find-SeElement -Driver $Driver -Name "OwnedBy"
                If (($Custodian -eq $Element.GetAttribute("value")) -or (($Custodian -eq "Enterprise Infrastructure" -and $Element.GetAttribute("value") -eq "Infrastructure, Enterprise"))) { #<----------
                    Write-Host "Finance Custodian" $Custodian "matches, no update" -ForegroundColor green
                }
                ElseIf (([string]::IsNullOrEmpty($Element.GetAttribute("value"))) -and ($Custodian = "Enterprise Infrastructure")) {
                    Write-Host "Finance Custodian tab empty, updating to $Custodian" -ForegroundColor blue
                    Send-SeKeys -Element $Element -Keys "Infrastructure, Enterprise"
                }
                Else {
                    $Element = $Element.Clear()
                    If ($Null -ne $Custodian) {
                        $Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasPrimaryUser"
                        Send-SeKeys -Element $Element -Keys $Custodian
                        Write-Host "Finance Custodian does not match, updating to $Custodian" -ForegroundColor blue
                    }                
                }
            }
            Catch {
                Write-Host "Cannot find Finance tab"
            } 
        } while ($Failed)
        Start-Sleep -Seconds 2

        <# This block finds the General tab button. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -TagName a
            Invoke-SeClick -Element $Element[41]
        } 
        Catch {
            Write-Host "Failed to click the General tab button." -ForegroundColor red
        }

        <# This block checks and verifies the hardware asset status. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -ClassName "k-ext-dropdown"
            If ($HardwareAssetStatus -eq $Element[4].GetAttribute("value")) {
                Write-Host "Hardware Asset Status" $HardwareAssetStatus "matches, no update" -ForegroundColor green
            }
            Else {
                $Element[4].Clear()
                Send-SeKeys -Element $Element[4] -Keys $HardwareAssetStatus
                $ElementValue = $Element[4].GetAttribute("value")
                Write-Host "Hardware Asset Status does not match, updating $ElementValue to $HardwareAssetStatus" -ForegroundColor blue
            }
        }
        Catch {
            Write-Host "Something went wrong with inputting $HardwareAssetStatus into Hardware Asset Status." -ForegroundColor red 
        }

        <# This block checks and verifies the location. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasLocation" 
            If ($LocationTag -eq $Element.GetAttribute("value")) {
                Write-Host "Location" $LocationTag "matches, no update" -ForegroundColor green
            }
            Else {
                $Element = $Element.Clear()
                $Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasLocation" 
                Send-SeKeys -Element $Element -Keys $LocationTag
                Write-Host "Location does not match, updating $ElementValue to $LocationTag" -ForegroundColor blue
            }
        }
        Catch {
            Write-Host "Something went wrong with inputting $LocationTag into Location." -ForegroundColor red 
        }

        <# This block checks and verifies the location details. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -Name "LocationDetails"
            $Element = $Element.Clear()
            If ($Null -ne $Office) {
                $Element = Find-SeElement -Driver $Driver -Name "LocationDetails"
                Send-SeKeys -Element $Element -Keys $Office
                Write-Host "Location Details has been updated to $Office" -ForegroundColor blue
            }
        }
        Catch {
            Write-Host "Something went wrong with inputting $Office into Location Details." -ForegroundColor red 
        }
        Start-Sleep -Seconds 2

        <# This block checks and inputs initials into the description box. #>
        Try {
            $Element = Find-SeElement -Driver $Driver -Name "Description"
            $descriptionContents = $Element.GetAttribute("value")
            $Element.Clear()
            Send-SeKeys -Element $Element -Keys $myInitials
            Send-SeKeys -Element $Element -Keys `n
            Send-SeKeys -Element $Element -Keys $descriptionContents
            Write-Host "Description has been updated with $myInitials" -ForegroundColor blue
        }
        Catch {
            Write-Host "Something went wrong with the description." -ForegroundColor red 
        }

        <# This block checks and verifies the front page custodian. #>
        do {
            Try {
                $Failed = $false
                $Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasPrimaryUser"
                If (($Custodian -eq $Element.GetAttribute("value")) -or ($Custodian -eq "Enterprise Infrastructure" -and $Element.GetAttribute("value") -eq "Infrastructure, Enterprise")) {
                    Write-Host "Custodian" $Custodian "matches, no update" -ForegroundColor green
                }
                ElseIf (([string]::IsNullOrEmpty($Element.GetAttribute("value"))) -and ($Custodian = "Enterprise Infrastructure")) {
                    Write-Host "Custodian tab empty, updating to $Custodian" -ForegroundColor blue
                    Send-SeKeys -Element $Element -Keys "Infrastructure, Enterprise"
                }
                Else {
                    $Element = $Element.Clear()
                    If ($Null -ne $Custodian) {
                        $Element = Find-SeElement -Driver $Driver -Name "Target_HardwareAssetHasPrimaryUser"
                        Send-SeKeys -Element $Element -Keys $Custodian
                        Write-Host "Custodian has been updated to $Custodian" -ForegroundColor blue
                    }
                }
            }
            Catch {
                Write-Host "Something went wrong with inputting $Custodian into Custodian." -ForegroundColor red 
                $Failed = $true
            }
        } while ($Failed)
        Start-Sleep -Seconds 2
    }
    ElseIf ($SearchItem -eq "exit") {
        Write-Host "Exit called -- ending the script." -ForegroundColor red
    }
    Else {
        Write-Host "Could not find $SearchItem in the spreadsheet." -ForegroundColor red
    }
}