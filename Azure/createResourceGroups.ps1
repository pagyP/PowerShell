<#Author       : Dean Cefola
# Creation Date: 08-15-2018
# Usage        : Create Multiple Resource Groups

#************************************************************************
# Date                         Version      Changes
#------------------------------------------------------------------------
# 08/15/2018                     1.0        Intial Version
#
#************************************************************************
#
#>

####################
#    Input Array   #
####################
$RGName = @(
    @{Name="bcprdrgpnetx001";Location='uksouth'} 
    @{Name="bcprdrgplawx001";Location='uksouth'} 
    @{Name="bcprdrgpbacx001";Location='uksouth'} 
    @{Name="bcprdrgptmpx001";Location='uksouth'} 
    
)


###############################
#    Create Resource Groups   #
###############################
foreach ($RG in $RGName) {
    
    if (-not(Get-AzResourceGroup -Name $RGName -ErrorAction SilentlyContinue)) {
        
    
    New-AzResourceGroup `
        -Name $RG.Name `
        -Location $RG.Location `
        -Tag @{"Cost Centre"="11302";Environment="Prod";Function="Networking";Application="Network"} `
    }
    else {
        Write-Host "Resource Group '$rg' already exists "
    }
}

