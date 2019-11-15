#Create dev/test resource groups and assign roles
Select-AzSubscription -SubscriptionId "entersubID"
$devVnetRGName = "bcdevrgpnetx001"
$tempAppRGName = "bcdevrgptmpx001"
$networkGroup = Get-AzureADGroup -SearchString "G-AzureNetworks"
$sysopsGroup = Get-AzureADGroup -SearchString "G-AzureSysops"
$location = "UK South"
#$subID = Get-Azsubscription

New-AzresourceGroup -name $devVnetRGName -location $location -Tag @{"Cost Centre"="11302";Environment="Production";Function="Resource Group";Application="Network";"Application Owner"="Networks";Directorate="Policy, Strategy and ICT"}

New-AzresourceGroup -name $tempAppRGName -location $location -Tag @{"Cost Centre"="11302";Environment="Production";Function="Resource Group";Application="Temporary Application";"Application Owner"="Sysops";Directorate="Policy, Strategy and ICT"}

New-AzroleAssignment -objectID $networkGroup.objectID -roledefinitionname "Network Contributor" -resourcegroupname $prodVnetRGName

New-AzroleAssignment -objectID $sysopsGroup.objectID -roledefinitionname "Virtual Machine Contributor" -resourcegroupname $tempAppRGName





#New-AzroleAssignment -objectID group.objectID -scope $subID.ID