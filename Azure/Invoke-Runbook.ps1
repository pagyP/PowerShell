#This script invokes an Azure runbook.  It passes the required parameters to the runbook
#TAG Variables - Update as Required
$environment = "Production"
$costCentre = "777"
$function = "Resource Group"
$Application = "Sysops"
$AppOwner = "Sysops"
$Directorate = "Policy, Strategy and ICT"
#The new Resource Group Name - Update as required
$RGName = "RG-Test-RG17"
#Use the correct group ID for the RBAC permission - the below is for G-AzureSysops
$RGGroupID = "<security group guid>"
#Use the correct subscription ID. Comment out the subscription you are NOT creating the resource group in
#The below is the production subscription
#$subID = "<subid>"
#The below is the Dev/Test subscription
#subID = "<subid>"
#Update the below with the required role to assign 
$role = "Contributor"
$automationAccountName = "testingauto"
$runbookName = "createResourceGroups"


$params = @{"ResourceGroupName"=$RGName;"RGGroupID"=$RGGroupID;
"costCentre"=$costCentre;"Environment"=$environment;"function"=$function;"Application"=$Application;"AppOwner"=$AppOwner;"Directorate"=$Directorate;"subID"=$subID;role=$role}



Start-AzAutomationRunbook –AutomationAccountName $automationAccountName `
 -Name $runbookName `
 -ResourceGroupName "bcprd0" `
  –Parameters $params