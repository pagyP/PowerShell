$CurrentDomain = Get-ADDomain

New-ADOrganizationalUnit -Name:"MSPName" -Path:"$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Customer" -Path:"$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator

New-ADOrganizationalUnit -Name:"Users" -Path:"OU=MSPName,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
#New-ADOrganizationalUnit -Name:"Workstations" -Path:"OU=MSPName,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Groups" -Path:"OU=MSPName,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
#New-ADOrganizationalUnit -Name:"Internal IT" -Path:"OU=MSPName,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator

New-ADOrganizationalUnit -Name:"Admin Accounts" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Service Accounts" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Servers" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Users" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator
New-ADOrganizationalUnit -Name:"Groups" -Path:"OU=Customer,$CurrentDomain" -ProtectedFromAccidentalDeletion:$true -Server:$CurrentDomain.PDCEmulator


New-ADGroup -GroupCategory:"Security" -GroupScope:"DomainLocal" -Name:"MSPName Management" -Path:"OU=Groups,OU=MSPName,$CurrentDomain" -SamAccountName:"Domain MSPName Management" -Server:$CurrentDomain.PDCEmulator
New-ADGroup -GroupCategory:"Security" -GroupScope:"Global" -Name:"Customer Admin Users" -Path:"OU=Groups,OU=Customer,$CurrentDomain" -SamAccountName:"Domain MSPName Finance" -Server:$CurrentDomain.PDCEmulator
New-ADGroup -GroupCategory:"Security" -GroupScope:"Global" -Name:"All Customer Users" -Path:"OU=Groups,OU=Customer,$CurrentDomain" -SamAccountName:"Domain MSPName Marketing" -Server:$CurrentDomain.PDCEmulator
#New-ADGroup -GroupCategory:"Security" -GroupScope:"Global" -Name:"Domain MSPName IT" -Path:"OU=Security Groups,OU=MSPName,$CurrentDomain" -SamAccountName:"Domain MSPName IT" -Server:$CurrentDomain.PDCEmulator

#To add in - enable AD recycle bin.  ACLS on OUs for MSPName management group and customer admin groups.  Sites and services. 
# DNS forwarders to mgmt domain. automate trust relationship and populate MSPName group?