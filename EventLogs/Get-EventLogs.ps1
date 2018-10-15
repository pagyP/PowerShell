
$servers = Get-Content -Path c:\servers.txt 
$date = (get-date).AddDays(-4)
#Get the critical and error level events from the Application and System logs
foreach ($server in $servers) {
Get-WinEvent -ComputerName $server @{logname='application','system';level=1,2; starttime = $date} -MaxEvents 10 | Select-Object -Property ID, Machinename, Message, logname | Export-Csv -Path c:\events.csv -NoTypeInformation }