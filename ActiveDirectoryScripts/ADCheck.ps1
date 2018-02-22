 #Written by Craig Dempsey, 19/03/2013
# ** NOTE ** This script will only work if the folder C:\admin\Scripts is present.
# ** NOTE ** This script will only work with Powershell v3.
# ** NOTE ** To adapt this to work with Powershell v2, you will need to change BPA cmdlets parameters, because the names of the parameters are different from v2 to v3.
# ** NOTE ** This script needs to be run on a domain controller.
 
#Import Module Best Practices for Powershell v2. In v3 the module gets automatically loaded.
Import-Module BestPractices
 
#Set Variables
$date = get-date -UFormat "%Y%m%d%H%M%S"
$date2 = get-date
$dcdiagcom = "dcdiag"
$dcdiaglog = "C:\Admin\Scripts\adchex\dcdiag$date.log"
$dcdiagargs = @('/a', '/c', '/v', "/f:$dcdiaglog ")
$repadmincom = "repadmin"
$repadminargs = @('/showrepl', '*', '/verbose', '/all', '/intersite')
$repadminlog =   "C:\admin\scripts\adchex\repl$date.log"
$ADbparesultcsv = "C:\Admin\Scripts\adchex\ADBpaResult$date.csv"
 
#Run the cmd commands calling the args.
&cmd /c $dcdiagcom $dcdiagargs
&cmd /c $repadmincom $repadminargs > $repadminlog
 
#Run the Best Practice Analayser
invoke-bpamodel -ModelId Microsoft/Windows/DirectoryServices
#Format the results
get-bparesult -ModelID Microsoft/Windows/DirectoryServices | Where { $_Severity -ne "Information" } | Set-BpaResult -Exclude $true| Export-CSV -Path $ADbparesultcsv
 
#Set email variables
$EmailFrom = "powershell@yourdomain.com"
$EmailTo = "whoeveruare@yourdomain.com"
$Subject = "AD CHEX!"
$Body = "Attached is a set of automated reports for your perusal. The reports contain a DCDiag report, a Repadmin report and Best Practice Analyser Report."
$SMTPServer = "YourSMTPserver"
 
#Email the log files.
Send-MailMessage -Subject $Subject -Body $body -SmtpServer $SMTPServer -Priority High -To $EmailTo -From $EmailFrom -Attachments $dcdiaglog, $repadminlog, $ADbparesultcsv