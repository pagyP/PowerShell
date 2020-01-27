# Variables for common values
$vNetrgName='RG_Networks_WUK'
$location='ukwest'
$recoveryVaultRG = "RG_Recovery_Vault_WUK"
$recoveryVaultName = "Vault1"

try {
    Get-AzureRmResourceGroup -Name $vNetrgName -ErrorAction Stop
} catch {
    New-AzureRmResourceGroup -Name $vNetrgName -Location $Location
}
# Create a resource group.
#New-AzResourceGroup -Name $rgName -Location $location

# Create a virtual network with a front-end subnet and back-end subnet.
$fesubnet = New-AzVirtualNetworkSubnetConfig -Name 'MySubnet-FrontEnd' -AddressPrefix '10.0.1.0/24'
$besubnet = New-AzVirtualNetworkSubnetConfig -Name 'MySubnet-BackEnd' -AddressPrefix '10.0.2.0/24'
New-AzVirtualNetwork -ResourceGroupName $rgName -Name 'MyVnet' -AddressPrefix '10.69.0.0/16'   -Location $location -Subnet $fesubnet, $besubnet

try {
    Get-AzureRmResourceGroup -Name $recoveryVaultRG -ErrorAction Stop
} catch {
    New-AzureRmResourceGroup -Name $recoveryVaultRG -Location $Location
}

New-AzRecoveryServicesVault - -Name $recoveryVaultName -Location $location
