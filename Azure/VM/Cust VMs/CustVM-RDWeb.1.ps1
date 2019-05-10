## Begin Variables    
#Global
$ResourceGroupName = "RG_Testing3"
$Location = "WestEurope"
$existingVnet = "devHubVnet"
$existingVnetResourceGroup = "RG_DevNetworkWE"
$avsetname = "Test-AS"
$random = Get-Random
$diagaccountname = "sardsdiag"+"$random"
$diskType = 'StandardSSD_LRS'
#$diskType = "Standard_LRS"
#$diskType = "Premium_LRS"
$VMName = "TestVM1"
$VMSize = "Standard_B2ms"
#To get a list of available VM sizes available in your chosen Azure region use 
# Get-AzVMSize -location 
#For example Get-AzVMImage -Location "West Europe"
$PublisherName = "MicrosoftWindowsServer"
$Offer = "WindowsServer"
$Sku = "2019-DataCenter-smalldisk"
#To get a list of publishers use Get-AzVMImagePublisher -location
#To get a list of offers use Get-AzVMImageOffer -Location location -PublisherName
#To get a listr of skus use Get-AzVMImageSku -Location -PublisherName -offer
##End Variables


# Create the resource group if needed
try {
   Get-azResourceGroup -Name $ResourceGroupName -ErrorAction Stop
} catch {
    New-azResourceGroup -Name $ResourceGroupName -Location $Location
}

#Storage account where the Initialise-VM script and UKRegion.xml are stored - UPDATE THIS!!
$fileUri = @("https://sapmpstorage.blob.core.windows.net/localesettings/Initialise-VM.ps1",
"https://sapmpstorage.blob.core.windows.net/localesettings/UKRegion.xml")

$Settings = @{"fileUris" = $fileUri};

$storageaccname = "sapmpstorage"
#Storage Account key not needed if the above files are in publically available storage
#$storagekey = "1234ABCD"
$ProtectedSettings = @{"storageAccountName" = $storageaccname;  "commandToExecute" = "powershell -ExecutionPolicy Unrestricted -File Initialise-VM.ps1"};

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
$Interface = New-azNetworkInterface -Name $InterfaceName -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $VNet.Subnets[2].Id 


#Define the VM Configuration
$vmConfig = New-azVMConfig -VMName $vmname -VMSize $VMSize -AvailabilitySetId $avset.Id | `
    Set-azVMOperatingSystem -Windows -ComputerName "$vmname" -Credential $credential -TimeZone 'GMT Standard Time' | `
    Set-azVMSourceImage -PublisherName $PublisherName -Offer $Offer -Skus $Sku -Version "latest" | `
    Add-azVMNetworkInterface -Id $interface.Id | `
    Set-azVMOSDisk -Name "$($vmname)-osdisk" -StorageAccountType $diskType -CreateOption FromImage | `
    #Add-azVMDataDisk -DiskSizeInGB 20 -Name "$($VMname)-datadisk" -Lun 0 -CreateOption Empty -StorageAccountType $diskType | `
    Set-azVMBootDiagnostics -Enable -ResourceGroupName $ResourceGroupName -StorageAccountName $diagaccountname

 #Create the VM in Azure
New-azVM -ResourceGroupName $ResourceGroupName -Location $Location -VM $vmConfig

#Apply Custom Script Extension which applies UK region settings to the VM
Set-azVMExtension -ResourceGroupName $ResourceGroupName -Location $Location -VMName $VMName -Name "localesettings" -Publisher "Microsoft.Compute" -ExtensionType "CustomScriptExtension"  -TypeHandlerVersion "1.9" -Settings $Settings -ProtectedSettings $ProtectedSettings 
#Apply Tags to the resource group
Set-AzResourceGroup -Name $ResourceGroupName -Tag @{ IaCMethod="PowerShell"; Customer="Lab"; Environment="Dev" }

$groups = Get-AzResourceGroup -Name $ResourceGroupName
foreach ($g in $groups)
{
    Get-AzResource -ResourceGroupName $g.ResourceGroupName | ForEach-Object {Set-AzResource -ResourceId $_.ResourceId -Tag $g.Tags -Force }
}