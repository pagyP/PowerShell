servers = Get-Content -Path c:\servers.txt
foreach ($server in $servers) {
    (Get-Hotfix -ComputerName $server | Sort-Object InstalledOn)[-1] | Export-Csv -Path c:\patchinginfo.csv -NoTypeInformation
}