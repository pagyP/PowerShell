# Variables for common values
$rgName='RG_Networks_NE'
$location='northeurope'


# Create a resource group.
New-AzResourceGroup -Name $rgName -Location $location

# Create a virtual network with a front-end subnet and back-end subnet.
$fesubnet = New-AzVirtualNetworkSubnetConfig -Name 'MySubnet-FrontEnd' -AddressPrefix '10.0.1.0/24'
$besubnet = New-AzVirtualNetworkSubnetConfig -Name 'MySubnet-BackEnd' -AddressPrefix '10.0.2.0/24'
New-AzVirtualNetwork -ResourceGroupName $rgName -Name 'MyVnet' -AddressPrefix '10.0.0.0/16'   -Location $location -Subnet $fesubnet, $besubnet