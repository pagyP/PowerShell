
#Set PowerShell Gallery as Trusted and import the modules pendingreboot and ImportExcel and add the AD Powershell feature to Windows
#https://www.powershellgallery.com/packages/PendingReboot/
#https://www.powershellgallery.com/packages/ImportExcel/5.3.4
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
Set-PSRepository -Name psgallery -InstallationPolicy Trusted
Install-module -Name PendingReboot 
Install-Module -Name ImportExcel 
Add-WindowsFeature -Name Rsat-AD-PowerShell
 
#Import from Active Directory any computer object which is running a server OS and is enabled and out to c:\computers.txt 
Get-ADComputer -Filter {(OperatingSystem -like "*Server*") -and (Enabled -eq $true)} -Properties OperatingSystem | Select-Object -ExpandProperty Name >c:\computers.txt
#Variables Section
$servers = Get-Content c:\computers.txt 
$date = (get-date).AddDays(-4)

#Use Test-PendingReboot module to query for pending reboots for Windows Updates and export to c:\serverinfo.xlsx
Test-PendingReboot -ComputerName $Servers -SkipPendingFileRenameOperationsCheck -SkipConfigurationManagerClientCheck | Export-Excel -WorksheetName pendingReboot -Path  c:\serverinfo.xlsx


#Here we get the most recent hotifx installed and the date on which it was installed
get-content c:\computers.txt | Where-Object {$_ -AND (Test-Connection $_ -Quiet)} | ForEach-Object { Get-Hotfix -computername $_ | Select-Object Csname,Description,HotFixID,InstalledBy,InstalledOn -Last 1 } | Export-Excel -worksheetname hotfix -path c:\serverinfo.xlsx 

#Get the most recent boot time of servers and export to c:\serverinfo.xlsx
Get-WmiObject win32_OperatingSystem -ComputerName $servers | Select-Object csname, @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}} | Export-Excel -WorksheetName lastboot -Path C:\serverinfo.xlsx

#Alternative event log collection
#foreach ($server in $servers)
#{Write-Host "Scanning the event log of: " -NoNewLine; Write-Host $server;
#Get-EventLog system -ComputerName $server -After (Get-Date).AddHours(-12) | where {($_.EntryType -Match "Error") -or ($_.EntryType -Match "Warning")} | select-object eventid, machinename, entrytype, message, timewritten | Export-csv -Path c:\serverinfo.csv;
#Get-EventLog application -ComputerName $server -After (Get-Date).AddHours(-12) | where {($_.EntryType -Match "Error") -or ($_.EntryType -Match "Warning")} | Export-Excel -WorksheetName AppLogErrorsandWarnings -Path c:\serverinfo.xlsx}

#Get the Warnings and Errors from the System event log
#Note sometimes get-winevent can throw errors if it can't parse the message text properly.  it still writes out the info to Excel but errors show in the console and
#the message text will say 'Cannot retrieve message text' in Excel.  This seems to happen if the event message has %% characters in it
foreach ($server in $servers){
    Get-Winevent -ComputerName $server -FilterHashtable @{
        logname = 'system'; level=1,2;starttime = $date} | Select-Object -Property ID, Machinename, Logname, Message | Export-Excel -WorksheetName SystemErrorsandWarnings -Path c:\serverinfo.xlsx 
}
#We have to do another pass for the application log as get-winevent only exports the info for the last server in the array if you specify
#more than one event log in the filter hash table
#Get the Warnings and Errors from the Application event log
foreach ($server in $servers){
    Get-Winevent -ComputerName $server -FilterHashtable @{
        logname = 'application'; level=1,2;starttime = $date} | Select-Object -Property ID, Machinename, Logname, Message | Export-Excel -worksheetname AppErrorsandWarnings -path c:\serverinfo.xlsx
    }



