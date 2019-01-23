Get-AzureRmNetworkSecurityGroup -ResourceGroupName "citrix-xd-d49dcd4b-2cc9-4c8e-860e-be6de69c239b-xvluj" -Name "Citrix-Deny-All-a3pgu-d49dcd4b-2cc9-4c8e-860e-be6de69c239b" | select SecurityRules -ExpandProperty SecurityRules


Get-AzureRmNetworkSecurityGroup -ResourceGroupName "citrix-xd-d49dcd4b-2cc9-4c8e-860e-be6de69c239b-xvluj" -Name "Citrix-Deny-All-a3pgu-d49dcd4b-2cc9-4c8e-860e-be6de69c239b" | select NetworkInterfaces



Get-AzureRmNetworkSecurityGroup -Name "Citrix-Deny-All-a3pgu-d49dcd4b-2cc9-4c8e-860e-be6de69c239b" -ResourceGroupName "citrix-xd-d49dcd4b-2cc9-4c8e-860e-be6de69c239b-xvluj" | Get-AzureRmNetworkSecurityRuleConfig | Select * 



Get-AzureRmNetworkInterface | where {$_.location -eq "uksouth"} | Select-Object -Property Name, @{Name="VMName";Expression = {$_.VirtualMachine.Id.tostring().substring($_.VirtualMachine.Id.tostring
    ().lastindexof('/')+1)}}



    Get-AzureRmNetworkSecurityGroup -ResourceGroupName RG-UKW-Citrix  -Name NSG-COVENS03     | Get-AzureRmNetworkSecurityRuleConfig | Select * | Export-Excel -WorksheetName NSG-COVENS03
  -Path C:\cove\info.xlsx