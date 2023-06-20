# PowerShell-Scripting
Here are my instructions below on how to use the PowerShell script I developed.

The script asks the user for an Asset Tag to which it will search the designated Excel spreadsheet for information. Then, it opens up the Microsoft Edge browser and automatically inputs in fields to search for the designated tag. After that, it will update the corresponding information obtained.

## Instructions to use:
In the Windows PowerShell console, 
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser # Use this if you are restricted by admin privileges.
Install-Module Selenium -Scope CurrentUser # This gives us the framework that allows us to manipulate the web browser.
```
It will ask you if you trust the repository -- go ahead and press [Y] to trust them.

You will need to have the [Microsoft Edge WebDriver](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/) installed.

Rename the application to **MicrosoftWebDriver.exe**. 

Place the .exe application in your Windows PowerShell Selenium Assemblies filepath 

*Example: C:\Users\CD4356\Documents\WindowsPowerShell\Modules\Selenium\3.0.1\assemblies

Now you can run the program by changing directory to where the script is downloaded in and running it.
```powershell
cd Documents # Where my script is in.
./script.ps1
```
Once the program has started, it will open up a Microsoft Edge browser; Put that browser onto one of your large monitors. It will then ask for an Asset Tag. Paste in a tag and it should automatically search it up for you in the system. After, it will log the information of the asset tag within the console. **Double check that the information is correct**. Once ready, save the page and repeat by pasting another Asset Tag. 

**You may exit the program whenever by typing "exit" or by hitting CTRL-C**.

Have fun!
### Documentation and References Used:
<ul>
  <li> https://github.com/adamdriscoll/selenium-powershell </li>
  <li> https://daniellange.wordpress.com/2009/12/18/searching-excel-in-powershell/ </li>
  <li> https://stackoverflow.com/questions/58100677/powershell-select-range-of-cells-from-excel-file-and-convert-to-csv </li>
  <li> https://www.powershellgallery.com/packages/Selenium/3.0.0/Content/Selenium.tests.ps1 </li>
  <li> https://devblogs.microsoft.com/scripting/powertip-new-lines-with-powershell/ </li>
</ul>
