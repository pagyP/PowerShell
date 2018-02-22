#############################
# Get AD Instantiation Date #
#############################
# Code from: http://blogs.technet.com/b/heyscriptingguy/archive/2012/01/05/how-to-find-active-directory-schema-update-history-by-using-powershell.aspx
write-output "Checking Active Directory Creation Date... " `r
write-output "Displaying AD partition creation information " `r

Import-Module ActiveDirectory
Get-ADObject -SearchBase (Get-ADForest).PartitionsContainer `
-LDAPFilter "(&(objectClass=crossRef)(systemFlags=3))" `
-Property dnsRoot,nETBIOSName,whenCreated | Sort-Object whenCreated | Format-Table dnsRoot,nETBIOSName,whenCreated -AutoSize