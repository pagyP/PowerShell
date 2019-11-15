#Create production resource groups and assign roles
Select-AzSubscription -SubscriptionId "enter subID"
$prodVnetRGName = "bcprdrgpnetx001"
$logAnalyticsRGName = "bcprdrgplawx001"
$backupASRRGName = "bcprdrgpbacx001"
$tempAppRGName = "bcprdrgptmpx001"
$networkGroup = Get-AzureADGroup -SearchString "G-AzureNetworks"
$sysopsGroup = Get-AzureADGroup -SearchString "G-AzureSysops"
$location = "UK South"
#$subID = Get-Azsubscription

New-AzresourceGroup -name $prodVnetRGName -location $location -Tag @{"Cost Centre"="11302";Environment="Production";Function="Resource Group";Application="Network";"Application Owner"="Networks";Directorate="Policy, Strategy and ICT"}
New-AzresourceGroup -name $logAnalyticsRGName -location $location -Tag @{"Cost Centre"="11302";Environment="Production";Function="Resource Group";Application="Azure Log Analytics";"Application Owner"="Sysops";Directorate="Policy, Strategy and ICT"}
New-AzresourceGroup -name $backupASRRGName -location $location -Tag @{"Cost Centre"="11302";Environment="Production";Function="Resource Group";Application="Azure Backup/ASR";"Application Owner"="Sysops";Directorate="Policy, Strategy and ICT"}
New-AzresourceGroup -name $tempAppRGName -location $location -Tag @{"Cost Centre"="11302";Environment="Production";Function="Resource Group";Application="Temporary Application";"Application Owner"="Sysops";Directorate="Policy, Strategy and ICT"}

New-AzroleAssignment -objectID $networkGroup.objectID -roledefinitionname "Network Contributor" -resourcegroupname $prodVnetRGName
New-AzroleAssignment -objectID $sysopsGroup.objectID -roledefinitionname "Log Analytics Contributor" -resourcegroupname $logAnalyticsRGName
New-AzroleAssignment -objectID $sysopsGroup.objectID -roledefinitionname "Automation Job Operator" -resourcegroupname $logAnalyticsRGName
New-AzroleAssignment -objectID $sysopsGroup.objectID -roledefinitionname "Automation Operator" -resourcegroupname $logAnalyticsRGName
New-AzroleAssignment -objectID $sysopsGroup.objectID -roledefinitionname "Backup Operator" -resourcegroupname $backupASRRGName
New-AzroleAssignment -objectID $sysopsGroup.objectID -roledefinitionname "Site Recovery Contributor" -resourcegroupname $backupASRRGName
New-AzroleAssignment -objectID $sysopsGroup.objectID -roledefinitionname "Virtual Machine Contributor" -resourcegroupname $tempAppRGName





#New-AzroleAssignment -objectID group.objectID -scope $subID.ID