param (
    [string] $secretName,
    [string] $subscriptionID,
    [string] $keyVaultName
)
    
        Add-Type -AssemblyName 'System.Web'
  
    
    
        $minLength = 25 ## characters
        $maxLength = 26 ## characters
        $length = Get-Random -Minimum $minLength -Maximum $maxLength
        $nonAlphaChars = 10
        $password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
        $secretvalue = ConvertTo-SecureString -String $password -AsPlainText -Force
        Select-AzSubscription  -SubscriptionId $subscriptionID
        Set-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName -SecretValue $secretvalue  > $null