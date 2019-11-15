$environment = "Production"
$costCentre = "777"
$function = "Resource Group"
$Application = "Sysops"
$AppOwner = "Sysops"
$Directorate = "Policy, Strategy and ICT"
$RGName = "RG-Test-RG15"
#Use the correct group ID for the RBAC permission - the below is for G-AzureSysops
$RGGroupID = "Enter Group ID"
$subID = "Enter SubID"
$automationAccountName = "testingauto"
$runbookName = "test2"


$params = @{"ResourceGroupName"=$RGName;"RGGroupID"=$RGGroupID;
"costCentre"=$costCentre;"Environment"=$environment;"function"=$function;"Application"=$Application;"AppOwner"=$AppOwner;"Directorate"=$Directorate;"subID"=$subID}



Start-AzAutomationRunbook –AutomationAccountName $automationAccountName `
 -Name $runbookName `
 -ResourceGroupName "bcprd0" `
  –Parameters $params