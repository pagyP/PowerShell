
$secretName = "enter the name to use for the secret"
$subscriptionID = "enter your subscription ID"
$keyVaultName = "enter your keyvault name"

#The below line requires Windows Powershell - it will not work in Powershell Core
Add-Type -AssemblyName 'System.Web'
$minLength = 25 ## characters
$maxLength = 26 ## characters
$length = Get-Random -Minimum $minLength -Maximum $maxLength
$nonAlphaChars = 10
$password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
$secretvalue = ConvertTo-SecureString -String $password -AsPlainText -Force

#Login to Azure subscription
Connect-AzAccount
#If the above connect-azaccount hangs your powershell window use Connect-Azaccount -UseDeviceAuthentication 

# Select subscription
Select-AzureRmSubscription  -SubscriptionId $subscriptionID

# Create Secret in Key Vault
Set-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName -SecretValue $secretvalue

# View Key
#(Get-AzKeyVaultSecret -vaultName "bcprdkeyxxxx001" -name "<name as appears in keyvault").SecretValueText


##azure devops pipeline version

# You can write your azure powershell scripts inline here. 
# You can also pass predefined and custom variables to this script using arguments
$secretName = ""
$subscriptionID = ""
$keyVaultName = ""

#The below line requires Windows Powershell - it will not work in Powershell Core
Add-Type -AssemblyName 'System.Web'
$minLength = 25 ## characters
$maxLength = 26 ## characters
$length = Get-Random -Minimum $minLength -Maximum $maxLength
$nonAlphaChars = 10
$password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
$secretvalue = ConvertTo-SecureString -String $password -AsPlainText -Force

#Login to Azure subscription
#Connect-AzAccount
#If the above connect-azaccount hangs your powershell window use Connect-Azaccount -UseDeviceAuthentication 

# Select subscription
Select-AzSubscription  -SubscriptionId $subscriptionID


# Create Secret in Key Vault.  Use output to null to prevent secret showing in pipeline log
Set-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName -SecretValue $secretvalue  > $null
