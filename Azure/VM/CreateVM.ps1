# 3 - Creating a new VM using the new managed disk we've created from a snapshot earlier.
 
$NewvirtualMachineName="OnurNewVM1"
$virtualNetworkName="Onur_Test-VNet"
$NewvirtualMachineSize="Standard_DS1"
$VNet = Get-AzureRmVirtualNetwork -Name $virtualNetworkName -ResourceGroupName $ResourceGroupName
$NIC = New-AzureRmNetworkInterface -Name ($NewvirtualMachineName.ToLower()+"_NIC") -ResourceGroupName $ResourceGroupName -Location $Location -SubnetId $VNet.Subnets[0].Id
 
$VirtualMachine = New-AzureRmVMConfig -VMName $NewvirtualMachineName -VMSize $NewvirtualMachineSize
$VirtualMachine = Set-AzureRmVMOSDisk -VM $VirtualMachine -ManagedDiskId $newOSDisk.Id -CreateOption Attach -Windows
$VirtualMachine = Add-AzureRmVMNetworkInterface -VM $VirtualMachine -Id $NIC.Id
$VirtualMachine = Set-AzureRmVMBootDiagnostics -VM $VirtualMachine -Disable
 
New-AzureRmVM -VM $VirtualMachine -ResourceGroupName $ResourceGroupName -Location $Location
