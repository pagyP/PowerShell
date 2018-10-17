###------------------------------###     
### Author : Biswajit Biswas-----###       
###--MCC, MCSA, MCTS, CCNA, SME--###     
###Email<bshwjt@gmail.com>-------###     
###------------------------------###     
###/////////..........\\\\\\\\\\\###     
###///////////.....\\\\\\\\\\\\\\### 
function Hotfixreport {  
    $computers = Get-ADComputer -Filter {(OperatingSystem -like "*Server*") -and (Enabled -eq $true)} -Properties OperatingSystem | select -ExpandProperty Name | Sort-Object
    #$computers = Get-Content C:\computers.txt    
    $ErrorActionPreference = 'Stop'    
    ForEach ($computer in $computers) {   
      
      try   
        {  
     
    Get-HotFix -cn $computer | Select-Object PSComputerName,HotFixID,InstalledOn,InstalledBy -last 3 | FT -AutoSize 
       
        }  
      
    catch   
      
        {  
    Add-content $computer -path "$env:USERPROFILE\Desktop\Notreachable_Servers.txt" 
        }   
    }  
      
    }  
    Hotfixreport > "$env:USERPROFILE\Desktop\Hotfixreport.txt" 