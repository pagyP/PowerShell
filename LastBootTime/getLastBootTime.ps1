
Get-CimInstance -ClassName win32_OperatingSystem | Select-Object csname, lastbootuptime

Get-WmiObject win32_OperatingSystem -ComputerName localhost | Select-Object csname, @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}} 
