$password = ConvertTo-SecureString -AsPlainText -String Password1234 -Force

Install-AddsForest -DomainName Corp.domain.cloud -SafeModeAdministratorPassword $password -DomainNetBiosName Corp -DomainMode 7 -ForestMode 7 -DataBasePath "%systemroot%\ntds" -LogPath "%systemroot%\ntds" -SysvolPath "%systemroot%\sysvol" -InstallDNS -Force 