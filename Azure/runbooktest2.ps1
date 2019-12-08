Param(
 [string]$ResourceGroupName,
 [string]$RGGroupID,
 [string]$costCentre,
 [string]$Environment,
 [string]$function,
 [string]$application,
 [string]$appowner,
 [string]$directorate,
 [string]$subID
)

$connectionName = "AzureRunAsConnection"

try
{
    # Get the connection "AzureRunAsConnection "
    $servicePrincipalConnection=Get-AutomationConnection -Name $connectionName  

    $tenantID = $servicePrincipalConnection.TenantId
    $applicationId = $servicePrincipalConnection.ApplicationId
    $certificateThumbprint = $servicePrincipalConnection.CertificateThumbprint

    "Logging in to Azure..."
    Connect-AzAccount `
        -ServicePrincipal `
        -TenantId $tenantID `
        -ApplicationId $applicationId `
        -CertificateThumbprint $certificateThumbprint
    
    Select-Azsubscription -subscriptionId $subID

    if (-not (Get-AzResourceGroup -Name $ResourceGroupName)) {

    New-AzResourceGroup -Name $ResourceGroupName -Location 'West Europe' -Tag @{"Cost Centre"=$costCentre;Environment=$Environment;Function=$function;Application=$application;"Application Owner"=$appowner;Directorate=$directorate}

    # Set the scope to the Resource Group created above
    $scope = (Get-AzResourceGroup -Name $ResourceGroupName).ResourceId

    # Assign Contributor role to the group
    New-AzRoleAssignment -ObjectId $RGGroupID -Scope $scope -RoleDefinitionName "Contributor"
    }
    else {
        write-host $ResourceGroupName "already exsists"
    }
}
catch {
   if (!$servicePrincipalConnection)
   {
      $ErrorMessage = "Connection $connectionName not found."
      throw $ErrorMessage
  } else{
      Write-Error -Message $_.Exception
      throw $_.Exception
  }
}
