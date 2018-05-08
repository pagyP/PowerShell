#Exports all GPOs to HTML
Get-GPO -all | % { Get-GPOReport -GUID $_.id -ReportType HTML -Path "d:\GPOExport\$($_.displayName).html" }
