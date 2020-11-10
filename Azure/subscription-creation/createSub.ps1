param (
    [parameter(Mandatory=$true)]
    [string]$EAAccount,
    [parameter(Mandatory=$true)]
    [string]$subscriptionName,
    [parameter(Mandatory=$true)]
    [string]$ManagementGroupName
)

$ErrorActionPreference='Stop'

try
{
	$OfferType='MS-AZR-0017P'

	# Get the Object ID of the VSTS SP running this script
	$upn = (Get-AzContext).Account.Id
	$spObjectId = (Get-AzADServicePrincipal -DisplayName $upn).ObjectId

	Write-Output "[INFO] UPN is [$upn] with object ID [$spObjectId]"

	Write-Output "[INFO] Checking subscription [$subscriptionName] already exist"

	$subscriptions = (Get-AzSubscription | Where-Object { $_.State -eq "Enabled" }).Name

	if($subscriptions -contains "$subscriptionName")
	{
		Write-Output "[WARN] Subscription [$subscriptionName] already exist, Exiting..."
	}
	else
	{
		Write-Output "[INFO] Subscription [$subscriptionName] does not exist, creating..."

		Write-Output "[INFO] Retrieving Enrollment account(s)"

		$eaAccounts = Get-AzEnrollmentAccount

		if(!($eaAccounts))
		{
			Write-Output "[ERROR] Enrollment account(s) not found. Exiting..."
			throw "[ERROR] Enrollment account(s) not found. Exiting..."
		}

		foreach($eaAcc in $eaAccounts)
		{
			if($eaAcc.PrincipalName.ToLower() -eq $EAAccount.ToLower())
			{
				$enrollmentName = $eaAcc.PrincipalName
				$enrollmentId = $eaAcc.ObjectId
			}
		}

		if(!($enrollmentName) -and ($enrollmentId))
		{
			throw "[ERROR] Enrollment account [$EAAccount] not found! Please check if the User/Service Principal has sufficient access. Exiting..."
		}

		#Now need to install the preview of the module

		Write-Output "[INFO] Updating Module..."
		Write-Warning "[WARN] Module versions have proven to be a problem in the past, leaving the historically required versions commented out in case of future trouble shooting..."

		Install-Module Az.Accounts 
		Install-Module Az.Subscription -RequiredVersion 0.7.2 -Force -AllowClobber
		Import-Module Az.Accounts 
		Import-Module Az.Subscription -Force

		Write-Output "[INFO] Creating the subscription [$subscriptionName]..."
		#Now create the subscription
		$newSub = New-AzSubscription -OfferType $OfferType -Name $subscriptionName -EnrollmentAccountObjectId $enrollmentId -OwnerObjectId $spObjectId

		if(!($newSub))
		{
			Write-Output "[ERROR] Failed to create the subscription [$subscriptionName]. Exiting..."
			throw $_.Exception.Message
		}

		Write-Output "[INFO] Subscription [$subscriptionName] created successfully with subscription Id : [$($newSub.Id)]"

		Write-Output "[INFO] Adding subscription [$subscriptionName] to the management group [$ManagementGroupName]"

		$retryCount = 0
		$maxRetries = 5
		$delayBeforeRetry = 120
		$completed = $false
		while(!($completed))
		{
			try
			{
				New-AzManagementGroupSubscription -GroupName $ManagementGroupName -SubscriptionId $newSub.Id
				$completed = $true
			}
			catch
			{
				if ($retryCount -ge $maxRetries)
				{
					Write-Output "[ERROR] Failed to move Subscription to Management Group [$ManagementGroupName] after 1 retry. Exiting."
					throw $_.Exception.Message
				}
				else
				{
					Write-Output "[WARN] Failed to move subscription to [$ManagementGroupName]. Will retry again after [$delayBeforeRetry] seconds..."
					Start-Sleep -Seconds $delayBeforeRetry
					$retryCount++
				}
			}
		}
	}
}
catch
{
	Write-Output "[ERROR] Failed to create/update subscription [$subscriptionName]. Exiting..."
	throw $_.Exception.Message
}
