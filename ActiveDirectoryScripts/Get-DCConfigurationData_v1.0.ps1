﻿<###################################################################################
Required information before executing:

    - Nothing, simply execute.

Execution:
    Get-DCConfigurationData_v1.0.ps1

This script will perform the following actions:

    - Checks to see if Domain_Discovery_Output folder exists.
        - If not, creates one under $Home\Documents.
    - Outputs a csv file to the Domain_Discovery_Output folder.
    - Gathers the following information about your domain controllers:
        - Server Name
        - Domain Name
        - Manufacturer
        - Model
        - Physical memory
        - OS caption
        - OS version
        - # of cores
        - CPU name
        - IP Address
        - DeviceID (hard drive letter)
        - HD size in GB
        - HD free space in GB
        - HD % free
        - NTDS and SYSVOL locations

Created by Theron Howton, MCS
           Fred Skaggs, formerly MCS
    07/26/2018

###################################################################################>

# Global variables
    $servers = (Get-ADGroupMember "domain controllers").name
    $date = Get-Date -UFormat "%m%d%Y-%H%M%S"
    $docs = "$home\Documents"

# Create output folder
If (!(Test-Path $docs\Domain_Discovery_Output)){
    New-Item -Path $docs\Domain_Discovery_Output -ItemType Directory
}
$output = "$docs\Domain_Discovery_Output"

Workflow Get-DCConfigurationData
{
    param(
            [string[]]$serverList
         )
             
    foreach -parallel ($server in $serverList)
    {
     
        sequence
        {

       
            # Get computer info
            $cs = Get-WmiObject -PSComputerName $server -Class 'Win32_ComputerSystem'

            # Get local disk info
            $ld = Get-WmiObject -PSComputerName $server `
                          -Query "select * from win32_logicaldisk where DriveType=3" | `
                            select DeviceID, SystemName, @{Name='Percent Free'; Expression={("{0:P2}" -f ($_.FreeSpace / $_.Size))}}, `
                            @{Name='Free Space GB'; Expression={[int]($_.FreeSpace/1GB)}}, `
                            @{Name='HD Size GB'; Expression={[int]($_.Size/1GB)}}

            # Get OS info
            $os = gwmi -PSComputerName $server -Class 'Win32_OperatingSystem'

            # Get processor info
            $proc = gwmi -PSComputerName $server -Class 'win32_processor'
            
            # Get IP addresses
            $ip = (Resolve-DNSName -Name $server -type A).ipaddress

            # Get NTDS/Sysvol folder locations
                $w32reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$server )
                $keypath1 = 'SYSTEM\\CurrentControlSet\\Services\\Netlogon\\Parameters'
                $RegKey1 = $w32reg.OpenSubKey($keypath1)
                $sysvol = $regkey1.GetValue('Sysvol')
                $keypath1 = 'SYSTEM\\CurrentControlSet\\Services\\NTDS\\Parameters'
                $RegKey2 = $w32reg.OpenSubKey($keypath1)
                $NTDS = $regkey2.GetValue('DSA Working Directory')
                      
            # Create custom property 
            $obj = New-Object -type PSObject -Property @{ 
                                                            "Server Name" = $server;
                                                            Domain = $cs.Domain;
                                                            Manufacturer = $cs.Manufacturer;
                                                            Model = $cs.Model;
                                                            "Physical Memory" = "{0:N2}" -f ($cs.TotalPhysicalMemory / 1GB);
                                                            CSName = $os.CSName;
                                                            "OS Caption" = $os.Caption;
                                                            "OS Version" = $os.Version;
                                                            "# of Cores" = $proc.NumberofCores;
                                                            "CPU Name" = $proc.Name;
                                                            Storage = $ld;
                                                            NTDS = $NTDS;
                                                            SYSVOL = $Sysvol;
                                                            IpAddress = $ip;
                                                        }

            # Return the custom object
            $obj
        }
    }
}

$dcs = Get-DCConfigurationData -serverList $servers #-PSPersist $True

$dcs | select -expandproperty Storage -property "Server Name", "Domain", "Manufacturer", "Model", "Physical Memory", "OS Caption", "OS Version", "# of Cores", "CPU Name", "IPAddress", "NTDS", "SYSVOL" `
     | select -Property "Server Name", "Domain", "Manufacturer", "Model", "Physical Memory", "OS Caption", "OS Version", "# of Cores", "CPU Name", "IPAddress", "DeviceID", "HD Size GB", "Free Space GB", "Percent Free", "NTDS", "SYSVOL" `
     | Export-Csv -Path "$output\All_DC_ConfigurationData_$($date).csv" -NoTypeInformation


Write-Host "Get-DCConfigurationData workflow complete." -ForegroundColor Green