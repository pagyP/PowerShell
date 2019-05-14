Workflow test {

    $vms = Get-AzurermVM | Where-Object {$_.tags['shutDown'] -eq "19:00"} 
    foreach -parallel ($vm in $vms) {
        #The following commands will be executed in parallel
        Stop-AZurermVM -Force   
    }
}
test