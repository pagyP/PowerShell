$count = Get-ADUser -searchbase "OU=Test,OU=Lab Users,DC=hidev,DC=local" -properties adminCount -filter * | where {$_.admincount -gt 0} >c:\admincount.csv
#$groups = Get-Adgroup -searchbase 
$count | select SamAccountname,distinguishedname,admincount | foreach-object {set-aduser -identity $_.samaccountname -replace @{admincount=0}}

