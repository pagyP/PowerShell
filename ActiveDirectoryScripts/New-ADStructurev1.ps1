$CurrentDomain = Get-ADDomain

New-ADOrganizationalUnit -Name:"Eduserv" -Path:"$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Customer" -Path:"$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator

New-ADOrganizationalUnit -Name:"Users" -Path:"OU=Eduserv,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
#New-ADOrganizationalUnit -Name:"Workstations" -Path:"OU=Eduserv,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Groups" -Path:"OU=Eduserv,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
#New-ADOrganizationalUnit -Name:"Internal IT" -Path:"OU=Eduserv,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator

New-ADOrganizationalUnit -Name:"Admin Accounts" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Service Accounts" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Servers" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Users" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Groups" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator


New-ADGroup -GroupCategory:"Security" -GroupScope:"DomainLocal" -Name:"Eduserv Management" -Path:"OU=Groups,OU=Eduserv,$CurrentDomain" -SamAccountName:"Domain Eduserv Management" -Server:$CurrentDomain.PDCEmulator
New-ADGroup -GroupCategory:"Security" -GroupScope:"Global" -Name:"Customer Admin Users" -Path:"OU=Groups,OU=Customer,$CurrentDomain" -SamAccountName:"Domain Eduserv Finance" -Server:$CurrentDomain.PDCEmulator
New-ADGroup -GroupCategory:"Security" -GroupScope:"Global" -Name:"All Customer Users" -Path:"OU=Groups,OU=Customer,$CurrentDomain" -SamAccountName:"Domain Eduserv Marketing" -Server:$CurrentDomain.PDCEmulator
#New-ADGroup -GroupCategory:"Security" -GroupScope:"Global" -Name:"Domain Eduserv IT" -Path:"OU=Security Groups,OU=Eduserv,$CurrentDomain" -SamAccountName:"Domain Eduserv IT" -Server:$CurrentDomain.PDCEmulator

#To add in - enable AD recycle bin.  ACLS on OUs for Eduserv management group and customer admin groups.  Sites and services. 
# DNS forwarders to mgmt domain. automate trust relationship and populate Eduserv group?