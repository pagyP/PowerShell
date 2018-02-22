Import-Module ActiveDirectory
$CSV=”C:\Scripts\NewComputerObjects.csv”
$OU=”OU=CengizTestOU,DC=kuskaya,DC=com”
Import-Csv -Path $CSV | ForEach-Object {New-ADComputer -Name $_.ComputerAccount -Path $OU -Enabled $True} 