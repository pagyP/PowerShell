<#-----------------------------------------------------------------------------
How to find AD schema update history using PowerShell
Ashley McGlone, Microsoft Premier Field Engineer
http://blogs.technet.com/b/ashleymcglone
December, 2011

This script reports on schema update and version history for Active Directory.
It requires the ActiveDirectory module to run.
It makes no changes to the environment.

UPDATED:
2013-03-12  Added Windows Server 2012, Exchange 2010 SP3 & 2013, Lync 2013

References for schema values:
http://support.microsoft.com/kb/556086?wa=wsignin1.0
http://social.technet.microsoft.com/wiki/contents/articles/2772.exchange-schema-versions-common-questions-answers.aspx
http://technet.microsoft.com/en-us/library/gg412822.aspx

LEGAL DISCLAIMER
This Sample Code is provided for the purpose of illustration only and is not
intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
nonexclusive, royalty-free right to use and modify the Sample Code and to
reproduce and distribute the object code form of the Sample Code, provided
that You agree: (i) to not use Our name, logo, or trademarks to market Your
software product in which the Sample Code is embedded; (ii) to include a valid
copyright notice on Your software product in which the Sample Code is embedded;
and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
against any claims or lawsuits, including attorneys’ fees, that arise or result
from the use or distribution of the Sample Code.
 
This posting is provided "AS IS" with no warranties, and confers no rights. Use
of included script samples are subject to the terms specified
at http://www.microsoft.com/info/cpyright.htm.
-----------------------------------------------------------------------------#>

Import-Module ActiveDirectory 

$schema = Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) `
    -SearchScope OneLevel -Filter * -Property objectClass, name, whenChanged,`
    whenCreated | Select-Object objectClass, name, whenCreated, whenChanged, `
    @{name="event";expression={($_.whenCreated).Date.ToShortDateString()}} | `
    Sort-Object whenCreated

"`nDetails of schema objects created by date:"
$schema | Format-Table objectClass, name, whenCreated, whenChanged `
    -GroupBy event -AutoSize

"`nCount of schema objects created by date:"
$schema | Group-Object event | Format-Table Count, Name, Group -AutoSize

#------------------------------------------------------------------------------

"`nForest domain creation dates:"
Get-ADObject -SearchBase (Get-ADForest).PartitionsContainer `
    -LDAPFilter "(&(objectClass=crossRef)(systemFlags=3))" `
    -Property dnsRoot, nETBIOSName, whenCreated |
  Sort-Object whenCreated |
  Format-Table dnsRoot, nETBIOSName, whenCreated -AutoSize

#------------------------------------------------------------------------------

$SchemaVersions = @()

$SchemaHashAD = @{
    13="Windows 2000 Server";
    30="Windows Server 2003 RTM";
    31="Windows Server 2003 R2";
    44="Windows Server 2008 RTM";
    47="Windows Server 2008 R2";
    56="Windows Server 2012 RTM"
    }

$SchemaPartition = (Get-ADRootDSE).NamingContexts | Where-Object {$_ -like "*Schema*"}
$SchemaVersionAD = (Get-ADObject $SchemaPartition -Property objectVersion).objectVersion
$SchemaVersions += 1 | Select-Object `
    @{name="Product";expression={"AD"}}, `
    @{name="Schema";expression={$SchemaVersionAD}}, `
    @{name="Version";expression={$SchemaHashAD.Item($SchemaVersionAD)}}

#------------------------------------------------------------------------------

$SchemaHashExchange = @{
    4397="Exchange Server 2000 RTM";
    4406="Exchange Server 2000 SP3";
    6870="Exchange Server 2003 RTM";
    6936="Exchange Server 2003 SP3";
    10628="Exchange Server 2007 RTM";
    10637="Exchange Server 2007 RTM";
    11116="Exchange 2007 SP1";
    14622="Exchange 2007 SP2 or Exchange 2010 RTM";
    14625="Exchange 2007 SP3";
    14726="Exchange 2010 SP1";
    14732="Exchange 2010 SP2";
    14734="Exchange 2010 SP3";
    15137="Exchange 2013 RTM"
    }

$SchemaPathExchange = "CN=ms-Exch-Schema-Version-Pt,$SchemaPartition"
If (Test-Path "AD:$SchemaPathExchange") {
    $SchemaVersionExchange = (Get-ADObject $SchemaPathExchange -Property rangeUpper).rangeUpper
} Else {
    $SchemaVersionExchange = 0
}

$SchemaVersions += 1 | Select-Object `
    @{name="Product";expression={"Exchange"}}, `
    @{name="Schema";expression={$SchemaVersionExchange}}, `
    @{name="Version";expression={$SchemaHashExchange.Item($SchemaVersionExchange)}}

#------------------------------------------------------------------------------

$SchemaHashLync = @{
    1006="LCS 2005";
    1007="OCS 2007 R1";
    1008="OCS 2007 R2";
    1100="Lync Server 2010";
    1150="Lync Server 2013"
    }

$SchemaPathLync = "CN=ms-RTC-SIP-SchemaVersion,$SchemaPartition"
If (Test-Path "AD:$SchemaPathLync") {
    $SchemaVersionLync = (Get-ADObject $SchemaPathLync -Property rangeUpper).rangeUpper
} Else {
    $SchemaVersionLync = 0
}

$SchemaVersions += 1 | Select-Object `
    @{name="Product";expression={"Lync"}}, `
    @{name="Schema";expression={$SchemaVersionLync}}, `
    @{name="Version";expression={$SchemaHashLync.Item($SchemaVersionLync)}}

#------------------------------------------------------------------------------

"`nKnown current schema version of products:"
$SchemaVersions | Format-Table * -AutoSize

#---------------------------------------------------------------------------sdg
