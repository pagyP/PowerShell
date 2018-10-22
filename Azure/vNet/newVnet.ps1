# Variables for common values
$rgName='RG_Networks_NE'
$location='northeurope'


# Create a resource group.
New-AzureRmResourceGroup -Name $rgName -Location $location

# Create a virtual network with a front-end subnet and back-end subnet.
$fesubnet = New-AzureRmVirtualNetworkSubnetConfig -Name 'MySubnet-FrontEnd' -AddressPrefix '10.0.1.0/24'
$besubnet = New-AzureRmVirtualNetworkSubnetConfig -Name 'MySubnet-BackEnd' -AddressPrefix '10.0.2.0/24'
New-AzureRmVirtualNetwork -ResourceGroupName $rgName -Name 'MyVnet' -AddressPrefix '10.0.0.0/16'   -Location $location -Subnet $fesubnet, $besubnet