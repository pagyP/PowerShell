## Begin Variables    
#Global
$ResourceGroupName = "RG-SAVE-Core-Services"
$Location = "WestEurope"
$existingVnet = "VNET-SAVE-Core"
$existingVnetResourceGroup = "RG-SAVE-Core-Network"
$avsetname = "AS-SAVE-Core-FileServices"

# Create the resource group if needed
#try {
  #  Get-AzureRmResourceGroup -Name $ResourceGroupName -ErrorAction Stop
#} catch {
 #   New-AzureRmResourceGroup -Name $ResourceGroupName -Location $Location
#}

#BootDiagStorage
$bootDiagsStorageName = "savecorediags"
$bootDiagsStorageResourceGroup = 'RG-SAVE-Core-Diags'

#Create the Availabiity Set
New-AzureRmAvailabilitySet -ResourceGroupName $ResourceGroupName -Location $Location -Name $avsetname -Sku Aligned -PlatformUpdateDomainCount 3 -PlatformFaultDomainCount 3
$avset = Get-AzureRMAvailabilitySet -ResourceGroupName $resourcegroupname -Name $avsetname




##Disk Storage Type
#$diskType = 'StandardSSD_LRS'
#diskType = "Standard_LRS"
$diskType = "Premium_LRS"

#Get the existing vnet
$VNet = Get-AzureRmVirtualNetwork -Name $existingVnet -ResourceGroupName $existingVnetResourceGroup

#Compute
$VMName = "SAVEFS001"
$VMSize = "Standard_DS2_v2"
# Create user object
$credential = Get-Credential -Message "Enter a username and password for the virtual machine."

#Define NIC Name
$InterfaceName = ($VMname.ToLower()+"-NIC")





#Create NIC, attach to subnet
#$PIp = New-AzureRmPublicIpAddress -Name $InterfaceName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Dynamic
$Interface = New-AzureRmNetworkInterface -Name $InterfaceName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $VNet.Subnets[5].Id 


#Define the VM Configuration
$vmConfig = New-AzureRmVMConfig -VMName $vmname -VMSize $VMSize -AvailabilitySetId $avset.Id | `
    Set-AzureRmVMOperatingSystem -Windows -ComputerName "$vmname" -Credential $credential -TimeZone 'GMT Standard Time' | `
    Set-AzureRmVMSourceImage -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2016-Datacenter" -Version "latest" | `
    Add-AzureRmVMNetworkInterface -Id $interface.Id | `
    Set-AzureRmVMOSDisk -Name "$($vmname)-osdisk" -StorageAccountType $diskType -CreateOption FromImage | `
    Add-AzureRmVMDataDisk -DiskSizeInGB 100 -Name "$($VMname)-datadisk" -Lun 0 -CreateOption Empty -StorageAccountType $diskType | `
    Set-AzureRmVMBootDiagnostics -Enable -ResourceGroupName $bootDiagsStorageResourceGroup -StorageAccountName $bootDiagsStorageName

 #Create the VM in Azure
New-AzureRmVM -ResourceGroupName $ResourceGroupName -Location $Location -VM $vmConfig

#Apply Custom Script Extension which applies UK region settings to the VM
#Set-AzureRmVMExtension -ResourceGroupName $ResourceGroupName -Location $Location -VMName $VMName -Name "localesettings" -Publisher "Microsoft.Compute" -ExtensionType "CustomScriptExtension"  -TypeHandlerVersion "1.9" -Settings $Settings -ProtectedSettings $ProtectedSettings 