# Step 1: Create a new resource group and key vault in the same location.
    # Fill in 'MyLocation', 'MySecureRG', and 'MySecureVault' with your values.
    # Use Get-AzureRmLocation to get available locations and use the DisplayName.
    # To use an existing resource group, comment out the line for New-AzureRmResourceGroup

    $Loc = 'MyLocation';
    $rgname = 'MySecureRG';
    $KeyVaultName = 'MySecureVault'; 
    New-AzureRmResourceGroup –Name $rgname –Location $Loc;
    New-AzureRmKeyVault -VaultName $KeyVaultName -ResourceGroupName $rgname -Location $Loc;
    $KeyVault = Get-AzureRmKeyVault -VaultName $KeyVaultName -ResourceGroupName $rgname;
    $KeyVaultResourceId = (Get-AzureRmKeyVault -VaultName $KeyVaultName -ResourceGroupName $rgname).ResourceId;
    $diskEncryptionKeyVaultUrl = (Get-AzureRmKeyVault -VaultName $KeyVaultName -ResourceGroupName $rgname).VaultUri;

#Step 2: Enable the vault for disk encryption.
    Set-AzureRmKeyVaultAccessPolicy -VaultName $KeyVaultName -ResourceGroupName $rgname -EnabledForDiskEncryption;

#Step 3: Create a new key in the key vault with the Add-AzureKeyVaultKey cmdlet.
    # Fill in 'MyKeyEncryptionKey' with your value.

    $keyEncryptionKeyName = 'MyKeyEncryptionKey';
    Add-AzureKeyVaultKey -VaultName $KeyVaultName -Name $keyEncryptionKeyName -Destination 'Software';
    $keyEncryptionKeyUrl = (Get-AzureKeyVaultKey -VaultName $KeyVaultName -Name $keyEncryptionKeyName).Key.kid;

#Step 4: Encrypt the disks of an existing IaaS VM
    # Fill in 'MySecureVM' with your value. 

    $VMName = 'MySecureVM';
    Set-AzureRmVMDiskEncryptionExtension -ResourceGroupName $rgname -VMName $vmName -DiskEncryptionKeyVaultUrl $diskEncryptionKeyVaultUrl -DiskEncryptionKeyVaultId $KeyVaultResourceId -KeyEncryptionKeyUrl $keyEncryptionKeyUrl -KeyEncryptionKeyVaultId $KeyVaultResourceId;