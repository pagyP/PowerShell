###########################
# Get Schema Update Dates #
###########################
# Code from: http://blogs.technet.com/b/heyscriptingguy/archive/2012/01/05/how-to-find-active-directory-schema-update-history-by-using-powershell.aspx
write-output "Reading all schema data... " `r
import-module activedirectory
$schema = Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) `
-SearchScope OneLevel -Filter * -Property objectClass, name, whenChanged,`
whenCreated | Select-Object objectClass, name, whenCreated, whenChanged, `
@{name="event";expression={($_.whenCreated).Date.ToShortDateString()}} | `
Sort-Object whenCreated

#"`nDetails of schema objects changed by date:"
#$schema | Format-Table objectClass, name, whenCreated, whenChanged `
#-GroupBy event -AutoSize

write-output "`nCount of schema objects changed by date:" `r
Write-output "This displays the approximate date each each schema update was performed." `r
$schema | Group-Object event | Format-Table Count,Name,Group –AutoSize