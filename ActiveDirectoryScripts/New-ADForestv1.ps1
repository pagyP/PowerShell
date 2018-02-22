#Step 1
Rename-Computer -NewName DC01
Restart-Computer -Force 

#Step 2
New-NetIPAddress –InterfaceIndex 12 –IPAddress 192.168.2.2 -PrefixLength 24
Set-DNSClientServerAddress –InterfaceIndex 12 -ServerAddresses 192.168.2.2

#Step 3
Install-WindowsFeature -Name AD-Domain-Services

#Step 4
$Password = ConvertTo-SecureString -AsPlainText -String !1Qwertyuiopüõ -Force
Install-ADDSForest -DomainName Corp.ViaMonstra.com -SafeModeAdministratorPassword $Password `
-DomainNetbiosName ViaMonstra -DomainMode Win2012R2 -ForestMode Win2012R2 -DatabasePath "%SYSTEMROOT%\NTDS" `
-LogPath "%SYSTEMROOT%\NTDS" -SysvolPath "%SYSTEMROOT%\SYSVOL" -NoRebootOnCompletion -InstallDns -Force

#Step 5
Restart-Computer -Force