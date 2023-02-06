<#
    .DESCRIPTION
        Automation function that syncs the assigned users/groups from an Enterprise App into a distribution group for use in transport rules and for communication.
		
    .NOTES
        AUTHOR: Michael Mardahl (github.com/mardahl)
        LASTEDIT: Mon feb 6th, 2023
#>

#region declarations
$tenantDomain = "xxxxxxx.onmicrosoft.com" #.onmicrosoft.com domain for exchange online connection
$graphVersion = "v1.0" #verison of Graph endpoint
$EnterpriseAppRegObjectId = "xx747a10-4x8f-4xxa-8xx8-599xxxx75exx" #The object Id of the Enterprise App that contains a list of assigned users/groups
$targetDistGroup = "xxxxxxx_dist" #Target Distribution group in exchange online
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
<#
"Please enable appropriate Microsoft Graph permissions to the system identity of this automation account. Otherwise, the runbook may fail..."
"The followign permissions can be given to the managed identity using this script: https://github.com/mardahl/PSBucket/blob/master/Add-MGraphMSIPermissions.ps1"
"Microsoft Graph : Group.ReadWrite.All, AppRoleAssignment.ReadWrite.All"
"Office 365 Exchange Online : Exchange.ManageAsApp"
"Azure AD RBAC role : Exchange Administrator"
#>

try
{
    "[INFO] Logging in to Azure with managed identity"
    Connect-AzAccount -Identity

	"[INFO] Acquire access token for Microsoft Graph"
	$token = (Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com').Token
	#$global:graphToken = @{Authorization="Bearer $token"}
	$global:graphToken = @{Authorization="Bearer $token";ConsistencyLevel="eventual"} #enables advanced queries

	"[INFO] Logging in to Exchange Online with managed identity"
	Connect-ExchangeOnline -ManagedIdentity -Organization $tenantDomain -ShowBanner:$false

}
catch {
    Write-Error -Message $_.Exception
    throw $_.Exception
}

#Get assigned users and groups for the App
"[INFO] Enumerating principals assigned to Enterprise App object id: $EnterpriseAppRegObjectId"
$assignees = Invoke-GraphRequest "/servicePrincipals/$EnterpriseAppRegObjectId/appRoleAssignedTo"

#Array to hold AAD object id's
$membersIdArray = @()

#Collect all pricipal ID's from assignee objects
foreach ($item in $assignees) {
    if($item.principalType -eq "Group"){
        #Get transitive members of group
        "[INFO] Getting transistive members of group $($item.PrincipalDisplayName) ($($item.principalId)"
        $members = Invoke-GraphRequest "/groups/$($item.principalId)/transitiveMembers/microsoft.graph.user?`$count=true&`$filter=accountEnabled eq true and mail ne null and userType eq 'Member'"
        $membersIdArray += $members.id
    } elseif($item.principalType -eq "User") {
        #item is a user and we already have the id
        $membersIdArray += $item.principalId
    }
}
$membersIdArray = $membersIdArray.Where({ $null -ne $_ }) #remove empty entries
"[INFO] Total number of 'appRoleAssignedTo' principals to sync with target distribution group ($targetDistGroup) is $($membersIdArray.count)"

"[INFO] Compiling lists for add/remove processing."
#Get list of existing distributiongroup members Object Ids
$distMembers = Get-DistributionGroupMember -Identity $targetDistGroup -Resultsize Unlimited | select ExternalDirectoryObjectId
#building list of users to remove
$toRemove = $distMembers | Where {$_.ExternalDirectoryObjectId -notin $membersIdArray}
"[INFO] $($toRemove.count) principals to remove"
#building list of users to add
$toAdd = $membersIdArray | Where {$_ -notin $distMembers.ExternalDirectoryObjectId}  
"[INFO] $($toAdd.count) principals to add"

#add members
if($toAdd.count -gt 0) {
    foreach ($objId in $toAdd){
        #validate that the member is an actual recipient object in Exchange Online
        $recipient = Get-Recipient -Identity $objId.trim() -ErrorAction SilentlyContinue | select Guid,PrimarySMTPAddress
        if ($recipient.count -ne 1) {
            "[ISSUE] The object id $objId is returning none or more than one recipient - could be a missing mailbox"
            $recipient
            continue
        }
        #add validated recipients
        try {
            Add-DistributionGroupMember -Identity $targetDistGroup -Member $recipient.Guid -ErrorAction Stop
            "[INFO] Added $($recipient.PrimarySMTPAddress)"
        } catch {
            $_
            "[WARNING] Unable to add $($recipient.PrimarySMTPAddress)"
        }
    }
} else {
    "[INFO] no members to be added"
}

#remove members
if($toRemove.count -gt 0) {
    foreach ($objId in $toRemove){
        try {
            $recipient = Get-Recipient -Identity $objId -ErrorAction SilentlyContinue | select Guid,PrimarySMTPAddress
            Remove-DistributionGroupMember -Identity $targetDistGroup -Member $recipient.PrimarySmtpAddress -Confirm:$false
            "[INFO] Removed $($recipient.PrimarySmtpAddress)"
        } catch {
            "[WARNING] Unable to remove $($member.PrimarySmtpAddress)"
        }
    }
} else {
    "[INFO] no members to be removed"
}

#enforce group restrictions
Set-DistributionGroup -Identity $targetDistGroup -MemberDepartRestriction closed -MemberJoinRestriction closed -WarningAction SilentlyContinue

#close the door nicely after execution
Disconnect-ExchangeOnline -Confirm:$false
#endregion execute
