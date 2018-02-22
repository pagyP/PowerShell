get-aduser -Properties Manager | Select-Object Name, @{n="ManagerName";e={(Get-ADUser -Identity $_.Manager -properties DisplayName).DisplayName}}
sort the list
take the managers name and group all users who have that manager set into a list
email that list to that manager