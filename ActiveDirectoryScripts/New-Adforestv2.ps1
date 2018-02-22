#Install AD DS,DNS and GPMC
$featureLogPath = "c:\poshlog\featurelog.txt"
#Change the variable domainname and netbios name to the customers name
$domainname = "customer.cust"
$netbiosName = "customer"
$Password = ConvertTo-SecureString -AsPlainText -String !1Qwertyuiop -Force
md c:\poshlog
start-job -Name addFeature -ScriptBlock {
Add-WindowsFeature -Name "ad-domain-services" -IncludeAllSubFeature -IncludeManagementTools
Add-WindowsFeature -Name "dns" -IncludeAllSubFeature -IncludeManagementTools
Add-WindowsFeature -Name "gpmc" -IncludeAllSubFeature -IncludeManagementTools }
Wait-Job -Name addFeature
Get-WindowsFeature | Where installed >>$featureLogPath
# Create New Forest, add Domain Controller
Import-Module ADDSDeployment
Install-ADDSForest -CreateDnsDelegation:$false `
-SafeModeAdministratorPassword $Password  `
-DatabasePath "C:\Windows\NTDS" `
-DomainMode "Win2012R2" `
-DomainName $domainname `
-DomainNetbiosName $netbiosName `
-ForestMode "Win2012R2" `
-InstallDns:$true `
-LogPath "C:\Windows\NTDS" `
-NoRebootOnCompletion:$true `
-SysvolPath "C:\Windows\SYSVOL" `
-Force:$true