#Peering for hub and spoke network design
#Variables Section
$hubVnetResourceGroup = "RG_hubVnet"
$hubVnetName = "hubprodVnet"
$spoke1VnetResourceGroup = "RG_spoke1vnet"
$spoke2VnetResourceGroup = "RG_spoke2Vnet"
$spoke1VnetName = "spoke1prodVnet"
$spoke2VnetName = "spoke2prodVnet"
$hubVnet = Get-AzureRmVirtualNetwork -Name $hubVnetName -ResourceGroupName $hubVnetResourceGroup 
$spoke1Vnet = Get-AzureRmVirtualNetwork -Name $spoke1VnetName -ResourceGroupName $spoke1VnetResourceGroup
$spoke2Vnet = Get-AzureRmVirtualNetwork -Name $spoke2VnetName -ResourceGroupName $spoke2VnetResourceGroup
#End Variables

#Add Hub to spoke1 peer and allow gateway transit through hub1
Add-AzureRmVirtualNetworkPeering -Name 'hubtospoke1peer' -VirtualNetwork $hubvnet -RemoteVirtualNetworkId $spoke1vnet.id -AllowForwardedTraffic  -AllowGatewayTransit 

#Add spoke 1 to hub and use hub 1 gateways
Add-AzureRmVirtualNetworkPeering -Name 'spoke1tohubpeer' -VirtualNetwork $spoke1vnet -RemoteVirtualNetworkId $hubVnet.id -AllowForwardedTraffic  -UseRemoteGateways 

#Add hub to spoke2 peer and allow gateway transit through hub
Add-AzureRmVirtualNetworkPeering -Name 'hubtospoke2peer' -VirtualNetwork $hubvnet -RemoteVirtualNetworkId $spoke2vnet.id -AllowForwardedTraffic  -AllowGatewayTransit

#Add spoke 2 to hub and use hub 1 gateways
Add-AzureRmVirtualNetworkPeering -Name 'spoke2tohubpeer' -VirtualNetwork $spoke2vnet -RemoteVirtualNetworkId $hubVnet.id -AllowForwardedTraffic  -UseRemoteGateways 





