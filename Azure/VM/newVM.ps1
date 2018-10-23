# Variables    
## Global
$ResourceGroupName = "RG_Webserver"
$Location = "NorthEurope"

## BootDiagStorage
#$bootDiagsStorageName = "sanebootdiagnostics"
#$bootDiagsStorageResourceGroup = 'RG_Networks_ne'


##Disk Storage Type
$diskType = 'StandardSSD_LRS'
#diskType = "Standard_LRS"
#diskType = "Premium_LRS"

## Get the existing vnet
$VNet = Get-AzureRmVirtualNetwork -Name myvnet -ResourceGroupName rg_networks_ne 

## Compute
$VMName = "web-01"
$ComputerName = "web-01"
$VMSize = "Standard_b1ms"
#$OSDiskName = $VMName + "-OSDisk"
# Create user object
$credential = Get-Credential -Message "Enter a username and password for the virtual machine."

#Define NIC Name
$InterfaceName = ($VMname.ToLower()+"-NIC")

# Resource Group
New-AzureRmResourceGroup -Name $ResourceGroupName -Location $Location



# Create NIC, attach to subnet
#$PIp = New-AzureRmPublicIpAddress -Name $InterfaceName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Dynamic
$Interface = New-AzureRmNetworkInterface -Name $InterfaceName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $VNet.Subnets[0].Id 


#Define the VM Configuration
$vmConfig = New-AzureRmVMConfig -VMName $vmname -VMSize $VMSize | `
    Set-AzureRmVMOperatingSystem -Windows -ComputerName "$vmname" -Credential $credential -TimeZone 'GMT Standard Time' | `
    Set-AzureRmVMSourceImage -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2016-Datacenter" -Version "latest" | `
    Add-AzureRmVMNetworkInterface -Id $interface.Id | `
    Set-AzureRmVMOSDisk -Name "$($vmname)-osdisk" -StorageAccountType $diskType -CreateOption FromImage | `
    Add-AzureRmVMDataDisk -DiskSizeInGB 20 -Name "$($VMname)-datadisk" -Lun 0 -CreateOption Empty -StorageAccountType $diskType | `
    Set-AzureRmVMBootDiagnostics -Enable -ResourceGroupName RG_Networks_NE -StorageAccountName sanebootdiagnostics

    ## Create the VM in Azure
New-AzureRmVM -ResourceGroupName $ResourceGroupName -Location $Location -VM $vmConfig