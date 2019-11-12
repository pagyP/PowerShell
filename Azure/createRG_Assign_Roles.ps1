$RGName = "RG-Test"
$Group = Get-AzureADGroup -SearchString "TestRbac"
$location = "west europe"
#$subID = Get-Azsubscription

New-AzresourceGroup -name $RGName -location $location
New-AzroleAssignment -objectID $group.objectID -roledefinitionname "Contributor" -resourcegroupname $RGName

#New-AzroleAssignment -objectID group.objectID -scope $subID.ID