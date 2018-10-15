#To run against servers in a domain review the below examples
#Get all enabled servers that are in the domain.
#$Servers = Get-ADComputer -Filter {(OperatingSystem -like "*Server*") -and (Enabled -eq $true)} -Properties OperatingSystem | select -ExpandProperty Name | Sort-Object
#At the moment the the below line pulls in a text file with server names.  The above query can be used instead in a domain environment
servers = Get-Content -Path c:\servers.txt
foreach ($server in $servers) {
    (Get-Hotfix -ComputerName $server | Sort-Object InstalledOn)[-1] | Export-Csv -Path c:\patchinginfo.csv -NoTypeInformation
}