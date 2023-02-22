# Enter the Firewall Name and RG Name as variables
$firewallName = "myFirewall"
$firewallResourceGroup = "myFirewallResourceGroup"

# Get the Azure Firewall
$azfw = Get-AzFirewall -Name $firewallName -ResourceGroupName $firewallResourceGroup
Write-Output "Processing Azure Firewall: $($azfw.id)..."

# Check if there is a management_ipconfigurations (indicates forced tunneling)
Write-Output "Checking for managementIpConfiguration (forced tunneling)..."
$fwRestApi = (Invoke-AzRestMethod -method GET -Path "$($azfw.id)`?api-version=2021-02-01").Content

if (($fwRestApi | ConvertFrom-Json).properties.managementIpConfiguration.id.count -gt 0)
{
    Write-Host -ForegroundColor Red "****ATTENTION****"
    Write-Output "Forced tunneling is enabled on this Azure Firewall. Stopping/Starting an Azure Firewall with Forced Tunneling enabled is not supported or possible. Attempting to start Azure Firewall with forced tunneling configured results in an HTTP 400 `"Bad Request`" error. Note that by continuing, the Azure Firewall will be stopped, and you will need to delete & recreate it from scratch. For more information about Forced Tunneling with Azure Firewall, please see https://docs.microsoft.com/en-us/azure/firewall/forced-tunneling"
    Write-Host -ForegroundColor Yellow "Please press ENTER to continue to stop the Azure Firewall, or CTRL+C to quit this script and do nothing."
    pause

    Write-Output "Stopping the Azure Firewall."
    Write-Output "For reference (to ease in firewall recreation), here is your Firewall config:"
    Write-Output $fwRestApi

    # Stop the Azure Firewall
    $azfw.Deallocate()
    Set-AzFirewall -AzureFirewall $azfw

    Write-Output "For reference and aid in resource recreation, here is your Firewall config:"
    Write-Output $fwRestApi
}
else
{
    # Output a Firewall restart script:
    Write-Output "`# Here's what info you will need to restart the Azure Firewall:"
    Write-Output "`$fwName = `"$($azfw.Name)`""
    Write-Output "`$rgName = `"$($azfw.ResourceGroupName)`""
    Write-Output "`$vnetName = `"$(($azfw.IpConfigurations[0].Subnet.id -split "/")[-3])`""
    Write-Output "`$vnetRg = `"$(($azfw.IpConfigurations[0].Subnet.id -split "/")[-7])`""
    $pipObjects = @()
    foreach ($i in 1..$($azfw.IpConfigurations.count))
    {
        Write-Output "`$pip$($i) = Get-AzPublicIpAddress -Name `"$(($azfw.ipConfigurations[$i-1].PublicIpAddress.id -split '/')[8])`" -ResourceGroupName `"$(($azfw.ipConfigurations[$i-1].PublicIpAddress.id -split '/')[4])`""
        $pipObjects += "`$pip$i"
    }
    Write-output ""
    Write-Output "`# Start the Azure firewall"
    Write-Output "`$azfw = Get-AzFirewall -Name `$fwName -ResourceGroupName `$rgName"
    Write-Output "`$vnet = Get-AzVirtualNetwork -Name `$vnetName -ResourceGroupName `$vnetRg"
    Write-Output "`$azfw.Allocate(`$vnet,@($($pipObjects -join ',')))"
    Write-Output "Set-AzFirewall -AzureFirewall `$azfw"

    # Stop the Azure Firewall
    $azfw.Deallocate()
    Set-AzFirewall -AzureFirewall $azfw
}
