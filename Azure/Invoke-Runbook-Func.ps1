[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$rgName,
    [Parameter(Mandatory=$true)]
    [string]$location,
    [Parameter(Mandatory=$true)]
    [string]$subId,
    [Parameter(Mandatory=$true)]
    [string]$environment,
    [Parameter(Mandatory=$true)]
    [string]$costCentre,
    [Parameter(Mandatory=$true)]
    [string]$function,
    [Parameter(Mandatory=$true)]
    [string]$Application,
    [Parameter(Mandatory=$true)]
    [string]$AppOwner,
    [Parameter(Mandatory=$true)]
    [string]$Directorate,
    [Parameter(Mandatory=$true)]
    [string]$RGGroupID,
    [Parameter(Mandatory=$true)]
    [string]$role
)
$automationAccountName = "testingauto"
$runbookName = "createResourceGroups"
#Example Usage
#Invoke-Runbook-Func.ps1 -rgName "RG-Test" -location "UK West" -subId "<subID> -environment "Production" -costCentre "11302"
# -function "Resource Group" -Application "Networks" -Appowner "Networks" -Directorate "Policy, Strategy and ICT" -RGGroupID "<group guid>" -role "Contributor"


$params = @{"ResourceGroupName"=$RGName;"RGGroupID"=$RGGroupID;
"costCentre"=$costCentre;"Environment"=$environment;"function"=$function;"Application"=$Application;"AppOwner"=$AppOwner;"Directorate"=$Directorate;"subID"=$subID;role=$role}

Start-AzAutomationRunbook –AutomationAccountName $automationAccountName `
 -Name $runbookName `
 -ResourceGroupName "bcprd0" `
  –Parameters $params