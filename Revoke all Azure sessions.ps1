
# Revoke all Azure sessions for all users based on UPN

Connect-MgGraph -Scopes User.RevokeSessions.All 

$users = Get-ADUser -Filter { UserPrincipalName -like "*@pol.falkenberg.se" -and Enabled -eq $true }

$uResults = $users | Select UserPrincipalName
$UPN = $uResults.UserPrincipalName
Write-Host "Found:"$users.Count"Accounts" -ForegroundColor Green

try {
    foreach ($UPN in $UPN) {
        Revoke-MgUserSignInSession -UserId $UPN -WhatIf
    }
    Write-Host "All sessions have been revoked" -ForegroundColor Green
}
catch {
    Write-Error $_
}
