$Domain = New-Object System.DirectoryServices.DirectoryEntry
$DomainSid = $Domain.objectSid
$RootDSE = New-Object System.DirectoryServices.DirectoryEntry(“LDAP://RootDSE”)
$RootDSE.UsePropertyCache = $false
$RootDSE.Put(“invalidateRidPool”, $DomainSid.Value)
$RootDSE.SetInfo()