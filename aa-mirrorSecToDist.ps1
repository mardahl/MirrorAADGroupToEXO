<#
    .DESCRIPTION
        Azure automation solution that mirrors membership of multiple security groups into corresponding distribution groups.
	Distributiongrouops that aren't preprovisioned will be created with the defined suffix.
		
    .NOTES
        AUTHOR: Michael Mardahl (github.com/mardahl)
        LASTEDIT: Nov 16, 2022
#>

#region declarations
$tenantDomain = "xxxxxx.onmicrosoft.com" #.onmicrosoft.com domain for exchange online connection
$graphVersion = "v1.0" #verison of Graph endpoint
$secGroupPrefix = "SecurityRollout_Wave" #prefix of the groups to mirror as Distribution groups
$distGroupSuffix = "_dist" #suffix added to the mirror groups. These are created if they don't exist

#endregion declarations

#region functions
function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory)]
        $query
    )

    $results = @()
    $url = "https://graph.microsoft.com/$graphVersion$query"

    while ($url) {
        $response = Invoke-RestMethod -Uri $url -Headers $graphToken -Method GET
        $results += $response.value

        # check for next link to continue paging
        $nextLink = ($response.'@odata.nextLink')
        if ($nextLink) {
            $url = $nextLink
        }
        else {
            $url = $null
        }

        # handle throttling
        if ($response.value.Count -eq 0 -and $response.statusCode -eq 429) {
            Write-Warning "Throttled! Waiting for 1 minute..."
            Start-Sleep -Seconds 60
        }
    }

    return $results
}
#endregion functions

#region execute
"Please enable appropriate Microsoft Graph permissions to the system identity of this automation account. Otherwise, the runbook may fail..."
"The followign permissions can be given to the managed identity using this script: https://github.com/mardahl/PSBucket/blob/master/Add-MGraphMSIPermissions.ps1"
"Microsoft Graph : Group.Read.All"
"Office 365 Exchange Online : Exchange.ManageAsApp"
"Azure AD RBAC role : Exchange Administrator (maybe less permissive 'Recipient Administrator' if possible)"

#connecting to Microsoft Graph and Exchange Online
try
{
    	"[INFO] Logging in to Azure with managed identity"
    	Connect-AzAccount -Identity

	"[INFO] Acquire access token for Microsoft Graph"
	$token = (Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com').Token
	$global:graphToken = @{Authorization="Bearer $token"}
	#$global:graphToken = @{Authorization="Bearer $token";ConsistencyLevel="eventual"} #enables advanced queries

	"[INFO] Logging in to Exchange Online with managed identity"
	Connect-ExchangeOnline -ManagedIdentity -Organization $tenantDomain -ShowBanner:$false

}
catch {
	Write-Error -Message $_.Exception
	throw $_.Exception
}

#Get Security Groups
$PSPersistPreference = $true
$SecurityGroups = Invoke-GraphRequest "/groups?`$filter=mailEnabled eq false and startsWith(displayName, '$secGroupPrefix')"
$PSPersistPreference = $false

#mirror each security group into a distribution group individually
foreach ($SecurityGroup in $SecurityGroups)
{
	$PSPersistPreference = $true
	$distGroupName = "$($SecurityGroup.displayName)$distGroupSuffix"
	#Get transitive members of security group
	$secMembers = Invoke-GraphRequest "/groups/$($SecurityGroup.id)/transitiveMembers"
	$PSPersistPreference = $false
	
	#find existing distribution groups and create a new one if none are found
	$distGroup = Get-DistributionGroup -Identity $distGroupName -ErrorAction SilentlyContinue

	if($distGroup){
		"[INFO] Existing group found ($distGroupName). Mirroring members."
		#building list of users to remove
		$distMembers = Get-DistributionGroupMember -Identity $distGroupName
		$toRemove = $distMembers | Where {$_.ExternalDirectoryObjectId -notin $secMembers.id}
		#building list of users to add
		$toAdd = $secMembers | Where {$_.id -notin $distMembers.ExternalDirectoryObjectId}  

		#add members
		if($toAdd -ne $null) {
			foreach ($member in $toAdd){
				#validate that the member is an actual recipient object in Exchange Online
				$recipient = Get-Recipient -Identity $member.Id -ErrorAction SilentlyContinue | select Guid,PrimarySMTPAddress
				if ($recipient -eq $null) {
					continue
				}
				#add validated recipients
				try {
					Add-DistributionGroupMember -Identity $distGroupName -Member $recipient.Guid -ErrorAction Stop
					"[INFO] Added $($recipient.PrimarySMTPAddress)"
				} catch {
					"[WARNING] Unable to add $($recipient.PrimarySMTPAddress)"
				}
			}
		} else {
			"[INFO] no new members to be added"
		}

		#remove members
		if($toRemove -ne $null) {
			foreach ($member in $toRemove){

				try {
					Remove-DistributionGroupMember -Identity $distGroupName -Member $member.PrimarySmtpAddress -Confirm:$false
					"[INFO] Removed $($member.PrimarySmtpAddress)"
				} catch {
					"[WARNING] Unable to remove $($member.PrimarySmtpAddress)"
				}
			}
		} else {
			"[INFO] no members to be removed"
		}

	} else {
		"[INFO] No group found ($distGroupName). Creating distribution group and mirroring members on next run."

		New-DistributionGroup -Name $distGroupName
	}
}
Disconnect-ExchangeOnline -Confirm:$false
#endregion execute
