
#Example below
#Get-AzureRmVM | where {$_.Tags['tag1'] -eq "value1" -and $_.Tags['tag2'] -eq "value2"}
#Get all VMs in the subscription via the tag name 'shutDown' and a tag value of '19:00'
Import-Module -Name AZ
Get-AzVM | Where-Object {$_.tags['shutDown'] -eq "19:00"} | Stop-AZVM -Force -AsJob

