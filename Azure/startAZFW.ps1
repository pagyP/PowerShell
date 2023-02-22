# Here's what info you will need to restart the Azure Firewall:
$fwName = "myFirewall"
$rgName = "myFirewallRg"
$vnetName = "myFirewallVnet"
$vnetRg = "myFirewallRg"
$pip1 = Get-AzPublicIpAddress -Name "myFirewall-IP01" -ResourceGroupName "myFirewallRg"
$pip2 = Get-AzPublicIpAddress -Name "myFirewall-IP02" -ResourceGroupName "myFirewallRg"
$pip3 = Get-AzPublicIpAddress -Name "myFirewall-IP03" -ResourceGroupName "myFirewallRg"
$pip4 = Get-AzPublicIpAddress -Name "myFirewall-IP04" -ResourceGroupName "myFirewallRg"
$pip5 = Get-AzPublicIpAddress -Name "myFirewall-IP05" -ResourceGroupName "myFirewallRg"
$pip6 = Get-AzPublicIpAddress -Name "myFirewall-IP06" -ResourceGroupName "myFirewallRg"
$pip7 = Get-AzPublicIpAddress -Name "myFirewall-IP07" -ResourceGroupName "myFirewallRg"
$pip8 = Get-AzPublicIpAddress -Name "myFirewall-IP08" -ResourceGroupName "myFirewallRg"
$pip9 = Get-AzPublicIpAddress -Name "myFirewall-IP09" -ResourceGroupName "myFirewallRg"
$pip10 = Get-AzPublicIpAddress -Name "myFirewall-IP10" -ResourceGroupName "myFirewallRg"
$pip11 = Get-AzPublicIpAddress -Name "myFirewall-IP11" -ResourceGroupName "myFirewallRg"
$pip12 = Get-AzPublicIpAddress -Name "myFirewall-IP12" -ResourceGroupName "myFirewallRg"
$pip13 = Get-AzPublicIpAddress -Name "myFirewall-IP13" -ResourceGroupName "myFirewallRg"
$pip14 = Get-AzPublicIpAddress -Name "myFirewall-IP14" -ResourceGroupName "myFirewallRg"
$pip15 = Get-AzPublicIpAddress -Name "myFirewall-IP15" -ResourceGroupName "myFirewallRg"
$pip16 = Get-AzPublicIpAddress -Name "myFirewall-IP16" -ResourceGroupName "myFirewallRg"
$pip17 = Get-AzPublicIpAddress -Name "myFirewall-IP17" -ResourceGroupName "myFirewallRg"
$pip18 = Get-AzPublicIpAddress -Name "myFirewall-IP18" -ResourceGroupName "myFirewallRg"
$pip19 = Get-AzPublicIpAddress -Name "myFirewall-IP19" -ResourceGroupName "myFirewallRg"
$pip20 = Get-AzPublicIpAddress -Name "myFirewall-IP20" -ResourceGroupName "myFirewallRg"

# Start the Azure firewall
$azfw = Get-AzFirewall -Name $fwName -ResourceGroupName $rgName
$vnet = Get-AzVirtualNetwork -Name $vnetName -ResourceGroupName $vnetRg
$azfw.Allocate($vnet,@($pip1,$pip2,$pip3,$pip4,$pip5,$pip6,$pip7,$pip8,$pip9,$pip10,$pip11,$pip12,$pip13,$pip14,$pip15,$pip16,$pip17,$pip18,$pip19,$pip20))
Set-AzFirewall -AzureFirewall $azfw
