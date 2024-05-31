
# Connect to Azure AD
Connect-AzureAD

# Get all enabled users with email addresses that end with @pol.domain.se
$users = Get-ADUser -Filter { UserPrincipalName -like "*@pol.falkenberg.se" -and Enabled -eq $true }

Write-Host "Found "$users.count" users." -ForegroundColor Green

$continue = Read-Host "Type Y to continue"

if ($continue -ne "Y") {
    return
}
# Sign out each user from all sessions in Azure AD
foreach ($user in $users) {
    Revoke-AzureADUserAllRefreshToken -ObjectId $user.UserPrincipalName -WhatIf
}
