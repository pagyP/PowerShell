#To run against servers in a domain review the below examples
#Get all enabled servers that are in the domain.
#$Servers = Get-ADComputer -Filter {(OperatingSystem -like "*Server*") -and (Enabled -eq $true)} -Properties OperatingSystem | select -ExpandProperty Name | Sort-Object
#At the moment the the below line pulls in a text file with server names.  The above query can be used instead in a domain environment
$servers = Get-Content -Path c:\servers.txt 
$date = (get-date).AddDays(-4)
#Get the critical and error level events from the Application and System logs
foreach ($server in $servers) {
Get-WinEvent -ComputerName $server @{logname='application','system';level=1,2; starttime = $date} -MaxEvents 10 | Select-Object -Property ID, Machinename, Message, logname | Export-Csv -Path c:\events.csv -NoTypeInformation }