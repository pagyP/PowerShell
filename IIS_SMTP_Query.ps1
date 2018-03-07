get-adcomputer -filter * | Select-Object dnshostname >c:\servers.txt
Get-WmiObject -Namespace "root\MicrosoftIISv2" -Class "IISSMTPServerSetting" -Filter "Name ='SmtpSvc/1'" -comp (Get-Content c:\servers.txt) | Select-Object smarthost,defaultdomain | export-csv c:\servers.csv -NoTypeInformation


