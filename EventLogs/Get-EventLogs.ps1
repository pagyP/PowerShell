#To run against servers in a domain review the below examples
#Get all enabled servers that are in the domain.
#$Servers = Get-ADComputer -Filter {(OperatingSystem -like "*Server*") -and (Enabled -eq $true)} -Properties OperatingSystem | select -ExpandProperty Name | Sort-Object
#At the moment the the below line pulls in a text file with server names.  The above query can be used instead in a domain environment
$servers = Get-Content -Path c:\servers.txt 
$date = (get-date).AddDays(-4)
#Get the critical and error level events from the Application and System logs
foreach ($server in $servers) {
Get-WinEvent -ComputerName $server @{logname='application','system';level=1,2; starttime = $date} -MaxEvents 10 | Select-Object -Property ID, Machinename, Message, logname | Export-Csv -Path c:\events.csv -NoTypeInformation }



$servers = get-content C:\servers1.txt
$days = Read-Host "History (Days)"
$BeginDate=[System.Management.ManagementDateTimeConverter]::ToDMTFDateTime((get-date).AddDays(-$days))
foreach ($computer in $servers) {
Get-WmiObject -ComputerName $computer  `
    -Query "SELECT ComputerName,Logfile,Type,TimeWritten,SourceName,Message,Category,EventCode,User `
    FROM Win32_NTLogEvent WHERE (logfile='Application') AND (type='Error') AND (TimeWritten > '$BeginDate')" | `
    SELECT ComputerName,Logfile,Type,@{name='TimeWritten';Expression={$_.ConvertToDateTime($_.TimeWritten)}},SourceName,Message,Category,EventCode,User | `
    Export-Csv "c:\Application-Errors.csv" }