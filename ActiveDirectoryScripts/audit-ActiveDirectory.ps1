<#
.SYNOPSIS

Audits a domain controller for compliance with the build guidelines

.DESCRIPTION

This is designed to audit the various settings to ensure that a domain controller has been built
to the build guidelines. There are 6 different tests that can be requested via the parameters
by default all tests are run.

.PARAMETER TestServer
	BOOLEAN Runs the following tests
		Audit-ServerBuildDetails 
		Audit-ServerNaming
		Audit-ServerDNSHostNameEntries
		Audit-ServerTimeSync

.PARAMETER TestDNS
	BOOLEAN Runs the following tests
		Audit-DNSServerConfiguration
		Audit-DNSZoneConfiguration

.PARAMETER TestNTDS
	BOOLEAN Runs the following tests
		Audit-ADBPath
		Audit-ADDBSize
		Audit-ADLogPath
		Audit-ADSYSVOLPath

.PARAMETER TestSites
	BOOLEAN Runs the following tests
		Audit-DefaultSiteChanges
		Audit-DefaultSiteLinkChanges

.PARAMETER TestRoles
	BOOLEAN Runs the following tests
		Audit-FSMORoles
		Audit-GCRoleHolder

.PARAMETER TestADHealth
	BOOLEAN Runs the following tests
		Audit-FunctionalLevels
		Audit-ReplicationErrors
		Audit-DCDiagResults

.PARAMETER LogFile
	STRING Specifies the filename for the report. By default this is stored in the current directory with filename
		Date-Time-Computername.htm

.EXAMPLE
audit-ActiveDirectory.ps1
	Run all tests and create output in current directory

.EXAMPLE
audit-ActiveDirectory.ps1 -TestADHealth -LogFile c:\adhealth.htm
	Run ADHealth test only and create out in adhealth.htm file
	
#>
[CmdletBinding()]
Param(
	[switch]$TestServer,
	[switch]$TestDNS,
	[switch]$TestNTDS,
	[switch]$TestSites,
	[switch]$TestRoles,
	[switch]$TestADHealth,
	[string]$LogFile = ".\" + [datetime]::Now.ToString("yyyMMdd-HHmm-") + $Env:COMPUTERNAME + ".htm"
)
$ADDBPath="E:\Windows\NTDS"
$ADLogPath="F:\Windows\NTDS"
$ADSYSVolPath="G:\Windows\NTDS"
$MaxNTDSDITSizeMB=1024 #1GB

function Audit-ADBPath () {
	$key = "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters"
	$valuename = "DSA Database file"
	$RunningPath = split-path((get-itemproperty $key -Name $valuename).$valuename) -parent
	if ($RunningPath.ToUpper() -eq $ADDBPath.ToUpper()) {
		Format-Output "TableRowOK" @("NTDS Database Path",$ADDBPath,$RunningPath)
	}
	else {
		Format-Output "TableRowError" @("NTDS Database Path",$ADDBPath,$RunningPath)
	}
}


function Audit-ADDBSize () {
	$key = "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters"
	$valuename = "DSA Database file"
	$CurrentDBSize = [int]((get-item ((get-itemproperty $key -Name $valuename).$valuename)).Length/1MB)
	if (($CurrentDBSize/$MaxNTDSDITSizeMB)*100 -lt 75) {
			Format-Output "TableRowOK" @("NTDS Database Size","<$MaxNTDSDITSizeMB",$CurrentDBSize)
	}
	elseif (($CurrentDBSize/$MaxNTDSDITSizeMB)*100 -lt 100) {
			Format-Output "TableRowWarn" @("NTDS Database Size","<$MaxNTDSDITSizeMB",$CurrentDBSize)
	}
	else {
			Format-Output "TableRowError" @("NTDS Database Size","<$MaxNTDSDITSizeMB",$CurrentDBSize)
	}
}


function Audit-ADLogPath () {
	$key = "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters"
	$valuename = "Database log files path"
	$RunningPath = (get-itemproperty $key -Name $valuename).$valuename
	if ($RunningPath.ToUpper() -eq $ADLogPath.ToUpper()) {
		Format-Output "TableRowOK" @("NTDS Log File Path",$ADLogPath,$RunningPath)
	}
	else {
		Format-Output "TableRowError" @("NTDS Database Path",$ADLogPath,$RunningPath)
	}
}


function Audit-ADSYSVOLPath () {
	$RunningPath = split-path((get-wmiobject -class win32_share | where-object {$_.Name -eq "SYSVOL"}).Path) -parent
	if ($RunningPath.ToUpper() -eq $ADSYSVolPath.ToUpper()) {
		Format-Output "TableRowOK" @("NTDS Database Path",$ADSYSVolPath,$RunningPath)
	}
	else {
		Format-Output "TableRowError" @("NTDS Database Path",$ADSYSVolPath,$RunningPath)
	}
}


function Audit-ServerDNSHostNameEntries () {
	$CountHostEntries=0
	$Result=nslookup  ($env:computername)
	for ($i=0;$i -lt $Result.length;$i++) {
		if($Result[$i].ToUpper() -match $env:computername.ToUpper()){
			$mkr=$i
		}
	}
	for ($i=$mkr+1;$i -lt $Result.length-1;$i++) {
		$IPs = $IPs + $Result[$i].Split(" ")[2] + "<br>"
		$CountHostEntries++
	}
	$IPs = $IPs.Remove($IPs.Length -4,4)
	if ($CountHostEntries -eq 1) {
		Format-Output "TableRowOK" @("DNS Server Hostname Entries","Single Entry",$IPs)
	}
	else {
		Format-Output "TableRowError" @("DNS Server Hostname Entries","Multiple Entries",$IPs)
	}
}


function Audit-FSMORoles () {
	$SchemaMaster = (Get-ADForest).SchemaMaster
	$DomainNamingMaster = (Get-ADForest).DomainNamingMaster
	$RIDMaster = (Get-ADDomain).RIDMaster
	$PDCEmulator = (Get-ADDomain).PDCEmulator
	$InfrastructureMaster = (Get-ADDomain).InfrastructureMaster
	$Site = (get-addomaincontroller).Site
	
	If ($SchemaMaster=$DomainNamingMaster) {
		Format-Output "TableRowOK" @("Schema Master and Domain Naming Master on Same Server", "Schema Master = Domain Naming Master", "Schema:$SchemaMaster<br>Domain:$DomainNamingMaster")
	}
	Else {	
		Format-Output "TableRowError" @("Schema Master and Domain Naming Master on Same Server", "Schema Master = Domain Naming Master", "Schema:$SchemaMaster<br>Domain:$DomainNamingMaster")
	}
	If ($PDCEmulator=$RIDMaster) {
		Format-Output "TableRowOK" @("PDC Emulator and RID Master on Same Server", "PDC Emulator = RID Master", "PDC Emulator:$PDCEmulator<br>RID Master:$RIDMaster")
	}
	Else {	
		Format-Output "TableRowOK" @("PDC Emulator and RID Master on Same Server", "PDC Emulator = RID Master", "PDC Emulator:$PDCEmulator<br>RID Master:$RIDMaster")
	}
}


function Audit-GCRoleHolder() {
	If ((get-addomaincontroller).IsGlobalCatalog -match "True") {
		Format-Output "TableRowOK" @("Server Is Global Catalog", "True", "True")
	}
	Else {
		Format-Output "TableRowError" @("Server Is Global Catalog", "True", "False")
	}	
}


function Audit-DCDiagResults () {
	$workfile = dcdiag /v
	$results = $workfile | foreach {$_ | `
				Where-Object {($_ -match "\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.") -or ($_ -match "Starting test")}}
	for ($i=0;$i -lt $results.length;$i++) {
		$DCDiagTest=$results[$i].split(" ")[8]
		$i++
		$DCDiagDetail=$results[$i].split(" ")[10]
		$DCDiagTestResult=$results[$i].split(" ")[11]
		if ($DCDiagTestResult -eq "passed") {
			Format-Output "TableRowOK" @("DCDiag Test: $DCDiagTest", "$DCDiagDetail", "$DCDiagTestResult")
		}
		else {
			Format-Output "TableRowError" @("DCDiag Test: $DCDiagTest", "$DCDiagDetail", "$DCDiagTestResult")
		}	
	}
}


function Audit-ReplicationErrors () {
	$workfile = repadmin /showrepl * /csv
	if ($workfile -eq $null) { return }
		$results = ConvertFrom-Csv -InputObject $workfile
		$results = $results | Select-Object @{Name="Source";Expression={$_."Source DSA"}}, `
			@{Name="Destination";Expression={$_."Destination DSA"}}, `
			@{Name="Partition";Expression={$_."Naming Context"}}, `
			@{Name="Failures";Expression={[int]$_."Number of Failures"}}, `
			@{Name="Failed";Expression={[datetime]::ParseExact($_."Last Failure Time","yyyy-MM-dd HH:mm:ss",$null)}}, `
			@{Name="Replicated";Expression={[datetime]::ParseExact($_."Last Success Time","yyyy-MM-dd HH:mm:ss",$null)}} `
			| Sort-Object -Property "Partition", "Source"
		$Partition=""
		$ErrorCount=0
	
	foreach ($Record in $results) {
		if (-not ($Record.Partition -eq $Partition)) {
			if (-not ($Partition -eq "")) {
				if ($ErrorCount -gt 0) {
					$FailedDate = $FailedDate.ToString("dd/MM/yyyy HH:mm:ss")
					Format-Output "TableRowError" @("Replication for<br>$Partition", "$Direction<br>$ErrorCount Failures", "$ReplicatedDate<br>$FailedDate")
				}
				elseif ($Partition -match "Server Down") {
					Format-Output "TableRowError" @("$Partition", "$Direction", "")
				}
				else{
					Format-Output "TableRowOK" @("Replication for<br>$Partition", "$Direction", "$ReplicatedDate")
				}
			}
			$Partition = $Record.Partition
			$Direction = $Record.Source + " -> " + $Record.Destination
			$ReplicatedDate = "Last Sync: " + $Record.Replicated.ToString("dd/MM/yyyy HH:mm:ss")
			$ErrorCount = $Record.Failures
			$FailedDate = $Record.Failed
		}
		else {
			$Direction = $Direction + "<br>" + $Record.Source + " -> " + $Record.Destination
			$ReplicatedDate = $ReplicatedDate + "<br>" + "Last Sync: " + $Record.Replicated.ToString("dd/MM/yyyy HH:mm:ss")
			$ErrorCount = $ErrorCount + $Record.Failures
			if (($FailedDate - $Record.Failed).Minutes -lt 0) {$FailedDate = $Record.Failed}
		}
	}
	if ($ErrorCount -gt 0) {
		$FailedDate = $FailedDate.ToString("dd/MM/yyyy HH:mm:ss")
		Format-Output  "TableRowError" @("Replication for<br>$Partition", "$Direction<br>$ErrorCount Failures", "$ReplicatedDate<br>$FailedDate")
	}
	elseif ($Partition -match "Server Down") {
		Format-Output "TableRowError" @("$Partition", "$Direction", "")
	}
	else {
		Format-Output  "TableRowOK" @("Replication for<br>$Partition", "$Direction", "$ReplicatedDate")
	}
}


function Audit-AllDCSReachable () {


}


function Audit-DefaultSiteChanges () {
  	$configNCDN = (Get-ADRootDSE).ConfigurationNamingContext
    $sitesContainerDN = ("CN=Sites," + $configNCDN)
	foreach ($site in Get-ADObject -Filter 'objectClass -eq "site"' -SearchBase $sitesContainerDN  -searchscope OneLevel ) {
		$sitelist=$sitelist + $site.Name + "<br>"
	}
	$sitelist =$sitelist.Remove($sitelist.Length -4,4)
	If ($sitelist.toupper() -match "DEFAULT-FIRST-NAME-SITE") {
	    Format-Output "TableRowError" @("Default Site Renamed", "Default-First-Name-Site", $sitelist)
	}
	else {
	    Format-Output "TableRowOK" @("Default Site Renamed", "Default-First-Name-Site", $sitelist)
	}
}


function Audit-DefaultSiteLinkChanges () {
	$configNCDN = (Get-ADRootDSE).ConfigurationNamingContext
	$siteLinkContainerDN = ("CN=IP,CN=Inter-Site Transports,CN=Sites," + $configNCDN)
	
	foreach ($siteLink in Get-ADObject -Filter 'objectclass -eq "siteLink"' -SearchBase $siteLinkContainerDN -SearchScope OneLevel ) {
		$Interval = (get-adobject $siteLink -properties replinterval).replInterval
		if ($Interval -eq 180) {$IntervalNotChanged=$true}
		$siteLinkList=$sitelist + $siteLink.Name + "&nbsp;" + $Interval + "<br>"
	}
	$siteLinklist =$siteLinklist.Remove($siteLinklist.Length -4,4)
	
	if ($IntervalNotChanged) {
		Format-Output "TableRowError" @("Site Replication Intervals Changed", "<> 180", $siteLinkList)
	}
	else {
		Format-Output "TableRowOK" @("Site Replication Intervals Changed", "<> 180", $siteLinkList)
	}
	if ($siteLinkList -match "DEFAULTIPSITELINK") {
		Format-Output "TableRowError" @("Default Site Link Renamed", "<> DEFAULTSITELINK", $siteLinkList)
	}
	else {
		Format-Output "TableRowOK" @("Default Site Link Renamed", "<> DEFAULTSITELINK", $siteLinkList)
	}	
}


function Audit-ServerBuildDetails () {
	$key = "HKLM:\Software\Cable and Wireless\Win2K8Base.Web.Hosting.Security.DC"
	if((Get-Item $key 2>$null| measure-object).Count -eq 1) {
		$key=$key -replace "Win2K8Base"," Win2K8Base"
		Format-Output "TableRowOK" @("Server Build Details",$key,"Found")
	}
	else {
		$key=$key -replace "Win2K8Base"," Win2K8Base"
		Format-Output "TableRowOK" @("Server Build Details",$key,"Not Found")
	}
}


function Audit-ServerNaming () {
	if (($env:COMPUTERNAME.Split("-")).Length -eq 4) {
		if (($env:COMPUTERNAME.Split("-"))[2].ToUpper() -eq "DC") {
			Format-Output "TableRowOK" @("Server Naming Standard", "XXX-YYY-DC-nnn", $env:computername)
		}
	}
	else {
		Format-Output "TableRowError" @("Server Naming Standard", "XXX-YYY-DC-nnn", $env:computername)
	}		
}


function Audit-DNSZoneConfiguration () {
$SecureUpdates="Secure"
$AllUpdates="Update"
$ForestLocation="AD-Forest"
$ReverseLookup="Rev"


	$Result=invoke-expression -command "dnscmd /enumzones"
	for ($i=0;$i -lt $Result.length;$i++) {
		if ($Result[$i].Replace([char]32," ") -match "Zone count") {
			$ZoneCount = $Result[$i].Replace([char]32," ").Replace("Zone count = ","")
		}
		if ($Result[$i].Replace([char]32," ") -match "Zone name") {
			$mkr = $i + 2
		}
	}
	for ($i=$mkr;$i -lt $mkr+$ZoneCount;$i++) {
		if ($Result[$i] -match "\.") {
			$Result[$i] = $Result[$i] -replace '\s+',',' 
			$t=$Result[$i].split(",")
			$ZoneName = $t[1]
			$Type = $t[2]
			$Storage =$t[3]
			if($t[4] -match $ReverseLookup) {
				$t[5]=$t[4]
				$t[4]="NoUpdate"
			}
			switch ($t[4]) {
				$SecureUpdates {
					$DynamicUpdate="Secure"
				}
				$AllUpdates {
					$DynamicUpdate="Update"
				}
				DEFAULT {
					$DynamicUpdate="No Updates"
				}
			}
			$ZoneText = "DNS Forward Lookup Zone - "
			if($t[5] -match $ReverseLookup) {
				$ReverseZone=$true
				$ZoneText = "DNS Reverse Lookup Zone - "
			}
			if ($Type -eq "Forwarder") {$ZoneText = "DNS Conditional Forwarder"}
			
			if (($Type -eq "Primary") -and ($DynamicUpdate -eq $SecureUpdates) -and ($Storage -eq $ForestLocation)) {
				Format-Output "TableRowOK" @("$ZoneText<br>$ZoneName", "Updates:$SecureUpdates<br>Storage:$ForestLocation<br>Type:$Type", "Updates:$DynamicUpdate<br>Storage:$Storage<br>Type:$Type")
			}
			elseif (($Type -eq "Forwarder") -and ($Storage -eq $ForestLocation)) {
				Format-Output "TableRowOK" @("$ZoneText<br>$ZoneName", "Updates:$SecureUpdates<br>Storage:$ForestLocation<br>Type:$Type", "Updates:$DynamicUpdate<br>Storage:$Storage<br>Type:$Type")
			}
			else {
				if (-not ($Type -eq "Cache")) {
					Format-Output "TableRowError" @("$ZoneText<br>$ZoneName", "Updates:$SecureUpdates<br>Storage:$ForestLocation<br>Type:$Type", "Updates:$DynamicUpdate<br>Storage:$Storage<br>Type:$Type")
				}
			}	
		}
	}
}


function Audit-DNSServerConfiguration () {
	$LogNothing=0
	$LogErrorsOnly=1
	$LogErrorsAndWarnings=2
	$LogAllEvents=7
	$LogSettings=@("No Logging","Log Errors","Log Errors And Warnings","","","","","Log All Events")
	
	$WantedLevel=$LogAllEvents
	if (-not ((get-itemProperty HKLM:\System\CurrentControlSet\services\DNS\Parameters) -match "EventLogLevel")) {
		$Loglevel = $WantedLevel
	}
	else {
		$Loglevel = (get-itemProperty HKLM:\System\CurrentControlSet\services\DNS\Parameters -Name EventLogLevel).EventLogLevel
	}
	if ($LogLevel -eq $WantedLevel) {
		Format-Output "TableRowOK" @("DNS Server Log Level",$LogSettings[$WantedLevel],$LogSettings[$Loglevel])
	}
	else {
		Format-Output "TableRowError" @("DNS Server Log Level",$LogSettings[$WantedLevel],$LogSettings[$Loglevel])
	}
	
}


function Audit-ServerTimeSync() {
	$NTPStatusLookup = @("No Sync","Holding","In Sync")
	$key="HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"
	$valuename="Type"
	$NTPType=(Get-ItemProperty $key -Name $valuename).$valuename
	$valuename="NtpServer"
	$NTPSources=((Get-ItemProperty $key -Name $valuename).$valuename)
	$NTPStatus = w32tm /query /status /verbose | Where-Object {($_ -match "Source:") -or ($_ -match "State Machine:")}
	$NTPCurrentSource = $NTPStatus[0] -replace "Source: "
	$NTPStatus = [int]($NTPStatus[1] -replace "[a-z:()]| ","")
	$NTPStatusText=$NTPStatusLookup[$NTPStatus]
	if (($NTPType -eq "NTP") -or ($NTPType -eq "AllSync")) {
		if ($NTPStatus -eq 2) {
			Format-Output "TableRowOK" @("NTP Configuration", "Source: NTP/AllSync<br>Status: Sync", "Source: $NTPType<br>Status: $NTPStatusText<br>Sourcelist: $NTPSources<br>Current Source: $NTPCurrentSource")
		}
		elseif ($NTPStatus -eq 1) {
			Format-Output "TableRowWarn" @("NTP Configuration", "Source: NTP/AllSync<br>Status: Sync", "Source: $NTPType<br>Status: $NTPStatusText<br>Sourcelist: $NTPSources<br>Current Source: $NTPCurrentSource")
		}
		else {
			Format-Output "TableRowError" @("NTP Configuration", "Source: NTP/AllSync<br>Status: Sync", "Source: $NTPType<br>Status: $NTPStatusText<br>Sourcelist: $NTPSources<br>Current Source: $NTPCurrentSource")
		}
	}
	else {
			Format-Output "TableRowError" @("NTP Configuration", "Source: NTP/AllSync", "Source: $NTPType<br>Sourcelist: $NTPSources<br>Current Source: $NTPCurrentSource")
	}
}


function Audit-FunctionalLevels () {
	$DomainFunctionalLevel=(get-ADDomain).DomainMode -replace "Domain",""
	$ForestFunctionalLevel=(get-ADForest).ForestMode -replace "Forest",""
	If ($DomainFunctionalLevel -eq "Windows2008R2") {
		Format-Output "TableRowOK" @("Domain Functional Level", "Windows2008R2", $DomainFunctionalLevel)
	}
	Elseif ($DomainFunctionalLevel -eq "Windows2000") {
		Format-Output "TableRowError" @("Domain Functional Level", "Windows2008R2", $DomainFunctionalLevel)
	}
	Else {
		Format-Output "TableRowWarn" @("Domain Functional Level", "Windows2008R2", $DomainFunctionalLevel)
	}
	If ($ForestFunctionalLevel -eq "Windows2008R2") {
		Format-Output "TableRowOK" @("Forest Functional Level", "Windows2008R2", $ForestFunctionalLevel)
	}
	Elseif ($ForestFunctionalLevel -eq "Windows2000") {
		Format-Output "TableRowError" @("Forest Functional Level", "Windows2008R2", $ForestFunctionalLevel)
	}
	Else {
		Format-Output "TableRowWarn" @("Forest Functional Level", "Windows2008R2", $ForestFunctionalLevel)
	}	

}


function Run-ServerChecks () {
	Format-Output "SectionStart" @("ServerChecks", "Server Configuration Tests")
	Audit-ServerBuildDetails 
	Audit-ServerNaming
	Audit-ServerDNSHostNameEntries
	Audit-ServerTimeSync
	Format-Output "SectionEnd"
}


function Run-DNSChecks () {
	Format-Output "SectionStart" @("DNSChecks", "DNS Configuration Tests")
	Audit-DNSServerConfiguration
	Audit-DNSZoneConfiguration
	Format-Output "SectionEnd"
}


function Run-NTDSChecks () {
	Format-Output "SectionStart" @("NTDSChecks", "AD NTDS Configuration Tests")
	Audit-ADBPath
	Audit-ADDBSize
	Audit-ADLogPath
	Audit-ADSYSVOLPath
	Format-Output "SectionEnd"
}


function Run-ADSiteChecks () {
	Format-Output "SectionStart" @("ADSiteChecks", "AD Site Configuration Tests")
	Audit-DefaultSiteChanges
	Audit-DefaultSiteLinkChanges
	Format-Output "SectionEnd"
}


function Run-ADRolesChecks () {
	Format-Output "SectionStart" @("FSMOChecks", "AD Role Holder Tests")
	Audit-FSMORoles
	Audit-GCRoleHolder
	Format-Output "SectionEnd"
}


function Run-ADHealthChecks () {
	Format-Output "SectionStart" @("ADHealthChecks", "AD Health Checks")
	Audit-FunctionalLevels
	Audit-ReplicationErrors
	Audit-AllDCSReachable
	Audit-DCDiagResults
	Format-Output "SectionEnd"
}


function Create-Header () {
	$Data=@"
<html>
<head>
<title>Active Directory Audit</Title>
<style type="text/css">
p.Normal, li.Normal, div.Normal, a.Normal, .Normal, Normal
	{
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:.0001pt;
	margin-left:0cm;
	font-size:11.0pt;
	font-family:"Calibri","sans-serif";
	}
tr.TableRowOK
	{
	background-color:#7CFF7F
	}
tr.TableRowError
	{
	background-color:#FF1500
	}
tr.TableRowWarn
	{
	background-color:#FF9707
	}
.h1, th.TableHeader
	{
	margin-top:24.0pt;
	margin-right:0cm;
	margin-bottom:0cm;
	margin-left:0cm;
	margin-bottom:.0001pt;
	line-height:115%;
	page-break-after:avoid;
	font-size:14.0pt;
	font-family:"Cambria","serif";
	color:#365F91;
	text-align:left;
	}
h2, tr.ColumnNames
	{
	margin-top:10.0pt;
	margin-right:0cm;
	margin-bottom:0cm;
	margin-left:0cm;
	margin-bottom:.0001pt;
	line-height:115%;
	page-break-after:avoid;
	font-size:13.0pt;
	font-family:"Cambria","serif";
	color:#4F81BD;
	background-color:#BDBDBD;
	}
p.Title, li.Title, div.Title, td.Title
	{
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:15.0pt;
	margin-left:0cm;
	border:none;
	padding:0cm;
	font-size:26.0pt;
	font-family:"Cambria","serif";
	color:#17365D;

	}
p.Subtitle, li.Subtitle, div.Subtitle
	{
	margin-top:0cm;
	margin-right:0cm;
	margin-bottom:10.0pt;
	margin-left:0cm;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Cambria","serif";
	color:#4F81BD;
	letter-spacing:.75pt;
	font-style:italic;
	}
td.col1
	{
	width:40%;
	}
td.col2, td.col3
	{
	width:30%;
	}
table.MainTable
	{
	width:100%;
	font-size:11.0pt;
	font-family:"Calibri","sans-serif";
	}
table.TOC
	{
	width:100%;
	font-size:11.0pt;
	font-family:"Calibri","sans-serif";
	}
table.SectionDetail
	{
	width:100%;
	font-size:11.0pt;
	font-family:"Calibri","sans-serif";
	}
</style>
</head>
<body>
<span class=Normal>
<table class=MainTable>
<tr>
<td class=Title>@ReportTitle</td>
</tr>
</table>


"@ + "`n"
return $Data.Replace("@ReportTitle",$env:COMPUTERNAME+ " generated on "+ (Get-Date).ToString("dd-MMM-yyyy") + " @" + (Get-Date).ToString("HH:mm"))
}


function Create-Footer () {
	$Data=$Data + @"
		<p align=center>--- End of Report ---</p>
		</span>
		</body>
		</html>
"@
	return $Data
}


function Format-Output ($Class, $Stuff) {
	switch -wildcard ($Class) {
		"SectionStart" {
			Update-TOC $Stuff[0] $Stuff[1]
			$HTML="<DIV ID=" + $Stuff[0] + " /><span class=h1>" + $Stuff[1] + "</span><Table class=SectionDetail>"
			$HTML=$HTML + "<TR class=columnnames><TD class=col1>Test</td><td class=col2>Required Value</td><td class=col3>Running Value</td></TR>"
	
		}
		"SectionEnd" {
			$HTML="</TABLE></DIV><br><br>"
		}
		"TableRow*" {
			$HTML = "<tr class=" + $Class + ">"
			foreach ($cell in $Stuff) {
				$HTML = $HTML + "<td>" + $cell + "</td>"
			}
			$HTML = $HTML + "</tr>"
		}
		DEFAULT {
			$HTML="<tr><td>" + $Stuff + "</td></tr>"
		}
	}
	$Script:HTML=$Script:HTML + $HTML
}


function Update-TOC ($Bookmark, $Description) {
	$Script:TOC=$Script:TOC + "<tr><td><a href=#$Bookmark class=Normal>$Description</a></td></tr>"
}



#---------------------------------------------------------------------------
#####################    Main Program Execution     ########################
#---------------------------------------------------------------------------
if ((Get-Module -listavailable | ?{$_ -match "ActiveDirectory"} | measure).count -eq 0) {
	Throw New-Object System.Management.Automation.RuntimeException("Unable to load ActiveDirectory module. You may need to install this feature and try again.")
}

Import-Module ActiveDirectory

$Header = Create-Header
$TOC="<BR><BR><table class=TOC>"
$HTML="</table><BR><BR>"

if (-not ($TestDNS -or $TestNTDS -or $TestRoles -or $TestServer -or $TestSites -or $TestADHealth)) {
	$TestDNS = $TestNTDS = $TestRoles = $TestServer = $TestSites = $TestADHealth =$true
}
if ($TestDNS) {	Run-DNSChecks }
if ($TestNTDS) { Run-NTDSChecks }
if ($TestRoles) { Run-ADRolesChecks }
if ($TestServer) { Run-ServerChecks }
if ($TestSites) { Run-ADSiteChecks }
if ($TestADHealth) { Run-ADHealthChecks }

$Footer = Create-Footer
$Header + $TOC + $HTML + $Footer | Out-File $LogFile

#---------------------------------------------------------------------------