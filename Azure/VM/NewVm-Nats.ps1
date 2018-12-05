## Begin Variables    
#Global
$ResourceGroupName = "RG_Test"
$Location = "WestEurope"
$existingVnet = "VNET-SAVE-Core"
$existingVnetResourceGroup = "RG-SAVE-Core-Network"
$avsetname = "Avset1"

# Create the resource group if needed
try {
    Get-AzureRmResourceGroup -Name $ResourceGroupName -ErrorAction Stop
} catch {
    New-AzureRmResourceGroup -Name $ResourceGroupName -Location $Location
}

#BootDiagStorage
$bootDiagsStorageName = "savecorediags "
$bootDiagsStorageResourceGroup = 'RG-SAVE-Core-Diags'

#Create the Availabiity Set
New-AzureRmAvailabilitySet -ResourceGroupName $ResourceGroupName -Location $Location -Name $avsetname 

#Storage account where the Initialise-VM script and UKRegion.xml are stored - UPDATE THIS!!
#$fileUri = @("https://saprodautomation.blob.core.windows.net/prodpowershelldsc/Initialise-VM.ps1",
#"https://saprodautomation.blob.core.windows.net/prodpowershelldsc/UKRegion.xml")

#$Settings = @{"fileUris" = $fileUri};

#$storageaccname = "saprodautomation"
#Storage Account key not needed if the above files are in publically available storage
#$storagekey = "1234ABCD"
#$ProtectedSettings = @{"storageAccountName" = $storageaccname;  "commandToExecute" = "powershell -ExecutionPolicy Unrestricted -File Initialise-VM.ps1"};

##Disk Storage Type
$diskType = 'StandardSSD_LRS'
#diskType = "Standard_LRS"
#diskType = "Premium_LRS"

#Get the existing vnet
$VNet = Get-AzureRmVirtualNetwork -Name $existingVnet -ResourceGroupName $existingVnetResourceGroup

#Compute
$VMName = "test01"
#$ComputerName = "web-01"
$VMSize = "Standard_b2ms"
#$OSDiskName = $VMName + "-OSDisk"
# Create user object
$credential = Get-Credential -Message "Enter a username and password for the virtual machine."

#Define NIC Name
$InterfaceName = ($VMname.ToLower()+"-NIC")

##End Variables



#Create NIC, attach to subnet
#$PIp = New-AzureRmPublicIpAddress -Name $InterfaceName -ResourceGroupName $ResourceGroupName -Location $Location -AllocationMethod Dynamic
$Interface = New-AzureRmNetworkInterface -Name $InterfaceName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $VNet.Subnets[3].Id 


#Define the VM Configuration
$vmConfig = New-AzureRmVMConfig -VMName $vmname -VMSize $VMSize | `
    Set-AzureRmVMOperatingSystem -Windows -ComputerName "$vmname" -Credential $credential -TimeZone 'GMT Standard Time' | `
    Set-AzureRmVMSourceImage -PublisherName "MicrosoftWindowsServer" -Offer "WindowsServer" -Skus "2016-Datacenter" -Version "latest" | `
    Add-AzureRmVMNetworkInterface -Id $interface.Id | `
    Set-AzureRmVMOSDisk -Name "$($vmname)-osdisk" -StorageAccountType $diskType -CreateOption FromImage | `
    Add-AzureRmVMDataDisk -DiskSizeInGB 20 -Name "$($VMname)-datadisk" -Lun 0 -CreateOption Empty -StorageAccountType $diskType | `
    Set-AzureRmVMBootDiagnostics -Enable -ResourceGroupName $bootDiagsStorageResourceGroup -StorageAccountName $bootDiagsStorageName

 #Create the VM in Azure
New-AzureRmVM -ResourceGroupName $ResourceGroupName -Location $Location  -VM $vmConfig

#Apply Custom Script Extension which applies UK region settings to the VM
#Set-AzureRmVMExtension -ResourceGroupName $ResourceGroupName -Location $Location -VMName $VMName -Name "localesettings" -Publisher "Microsoft.Compute" -ExtensionType "CustomScriptExtension"  -TypeHandlerVersion "1.9" -Settings $Settings -ProtectedSettings $ProtectedSettings 