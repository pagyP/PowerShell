## Begin Variables    
#Global
$ResourceGroupName = "RG-RDS"
$Location = "WestEurope"
$existingVnet = "core"
$existingVnetResourceGroup = "core"
$avsetname = "AS-SAVE-Core-RDSW"
$random = Get-Random
$diagaccountname = "sardsdiag"+"$random"
$diskType = 'StandardSSD_LRS'
#$diskType = "Standard_LRS"
#$diskType = "Premium_LRS"
$VMName = "SAVERDW001"
$VMSize = "Standard_DS2_v2"


# Create the resource group if needed
try {
   Get-azResourceGroup -Name $ResourceGroupName -ErrorAction Stop
} catch {
    New-azResourceGroup -Name $ResourceGroupName -Location $Location
}


#BootDiagStorage
New-AzStorageAccount -StorageAccountName $diagaccountname -ResourceGroupName $ResourceGroupName -Location $location -AccessTier hot -kind storagev2 -skuname Standard_LRS 
#$bootDiagsStorageName = "savecorediags "
#$bootDiagsStorageResourceGroup = 'RG-SAVE-Core-Diags'

#Create the Availabiity Set
New-azAvailabilitySet -ResourceGroupName $ResourceGroupName -Location $Location -Name $avsetname -Sku Aligned -PlatformUpdateDomainCount 3 -PlatformFaultDomainCount 3
#Now get the availability set so we can use it laterin the VM Config
$avset = Get-azAvailabilitySet -ResourceGroupName $resourcegroupname -Name $avsetname

#Get the existing vnet
$VNet = Get-azVirtualNetwork -Name $existingVnet -ResourceGroupName $existingVnetResourceGroup


# Create local admin account on Windows VM
$credential = Get-Credential -Message "Enter a username and password for the virtual machine."

#Define NIC Name
$InterfaceName = ($VMname.ToLower()+"-NIC")

#Create NIC, attach to subnet
#$PIp = New-azPublicIpAddress -Name $InterfaceName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Dynamic
$Interface = New-azNetworkInterface -Name $InterfaceName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $VNet.Subnets[1].Id 


#Define the VM Configuration
$vmConfig = New-azVMConfig -VMName $vmname -VMSize $VMSize -AvailabilitySetId $avset.Id | `
    Set-azVMOperatingSystem -Windows -ComputerName "$vmname" -Credential $credential -TimeZone 'GMT Standard Time' | `
    Set-azVMSourceImage -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2016-Datacenter" -Version "latest" | `
    Add-azVMNetworkInterface -Id $interface.Id | `
    Set-azVMOSDisk -Name "$($vmname)-osdisk" -StorageAccountType $diskType -CreateOption FromImage | `
    #Add-azVMDataDisk -DiskSizeInGB 20 -Name "$($VMname)-datadisk" -Lun 0 -CreateOption Empty -StorageAccountType $diskType | `
    Set-azVMBootDiagnostics -Enable -ResourceGroupName $ResourceGroupName -StorageAccountName $diagaccountname

 #Create the VM in Azure
New-azVM -ResourceGroupName $ResourceGroupName -Location $Location -VM $vmConfig

#Apply Custom Script Extension which applies UK region settings to the VM
#Set-azVMExtension -ResourceGroupName $ResourceGroupName -Location $Location -VMName $VMName -Name "localesettings" -Publisher "Microsoft.Compute" -ExtensionType "CustomScriptExtension"  -TypeHandlerVersion "1.9" -Settings $Settings -ProtectedSettings $ProtectedSettings 