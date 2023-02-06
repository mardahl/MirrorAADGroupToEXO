# Mirror AAD groups with Exchange Online groups
### aa-mirrorSecToDist.ps1
Azure Automation script that mirrors members of multiple Security Groups in Azure AD with members of corresponding Exchange Online Distribution Groups.

Can auto create the exchange online groups if missing. Uses Prefix to identify source groups, and suffix to identify target groups

# Sync Enterprise App assigned users/groups with Exchange Online distribution group
### aa-mirrorAppAssignedToDist.ps1
Azure automation script to keep a list of assigned principals for an Enterprise App in sync with a Distribution group.
