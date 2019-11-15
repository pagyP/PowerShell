$params = @{"ResourceGroupName"="enggdevsoutheastasia124";"RGGroupID"="f63da857-0a91-4f89-94b7-a97a9b9809c5"}
$automationAccountName = "testingauto"
$runbookName = "testagain"


Start-AzAutomationRunbook –AutomationAccountName $automationAccountName `
 -Name $runbookName `
 -ResourceGroupName "bcprd0" `
  –Parameters $params