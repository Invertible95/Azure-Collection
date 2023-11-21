<#
.SYNOPSIS
This PowerShell script performs the following tasks:
    - Connects to Microsoft Graph to retrieve information about Azure AD applications and their associated secrets.
    - Generates a report on secrets nearing expiry.
    - Requires appropriate permissions (Cloud Application Administrator role) to access Azure.
    - Checks for the presence of required modules (Microsoft.Graph.Authentication and ImportExcel) and installs them if necessary.
    - Exports a report to an Excel file with a timestamp, including application details, secret IDs, and expiry dates.

.MODULE INSTALLATION CHECK
To ensure the necessary modules are available, the script dynamically checks and installs them:

    $modulesToCheck = @("Microsoft.Graph.Authentication", "ImportExcel")

    foreach ($module in $modulesToCheck) {
        if (-not (Get-Module -Name $module -ListAvailable)) {
            Install-Module -Name $module -Force -AllowClobber
        }
    }
#>
# Import the required modules
Import-Module Microsoft.Graph.Authentication
Import-Module ImportExcel

# Connect to Microsoft Graph
Connect-AzAccount -Credential (Get-Credential)
Connect-MgGraph -Scopes "Application.ReadWrite.All"

$allApplications = @()
$nextLink = $null

do {
    # Retrieve paginated results of all applications
    $applicationsPage = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications\$($nextLink -replace '\?', '&')"
    
    $allApplications += $applicationsPage.Value
    $nextLink = $applicationsPage.'@odata.nextLink'
} while ($nextLink)

$retArr = @()

foreach ($Application in $allApplications) {
    $SecretURI = "https://graph.microsoft.com/v1.0/applications/$($Application.id)/passwordCredentials"
    try {
        # Retrieve secrets for each application
        $FindSecretDates = Invoke-MgGraphRequest -Method GET -Uri $SecretURI
    }
    catch {
        Write-Host "$_.ErrorDetails"
        continue
    }
        
    foreach ($expireDate in $FindSecretDates.Value) {
        # Process each secret and check expiry
        $expiryDateTime = [datetime]$expireDate.endDateTime
        $expiryDate = $expiryDateTime.Date
        
        if ($expiryDate -ne $null) {
            $daysUntilExpiry = ($expiryDate - (Get-Date).Date).Days
            if ($daysUntilExpiry -gt 30) {
                # Create a custom object for secrets nearing expiry
                $myObject = [PSCustomObject]@{
                    ApplicationName = $($Application.displayName)
                    "Object/app ID" = $($Application.id)
                    "Secret ID"     = $($FindSecretDates.value.keyId)
                    ExpiryDate      = $($expiryDate.ToString("yyyy-MM-dd"))
                }
                $retArr += $myObject
            } 
        } 
    }
}   

# Export the collected data to an Excel file with a timestamp
$excelDate = (Get-Date -Format "yyyy-MM-dd")
$retArr | Export-Excel -Path "C:\temp\Report $excelDate.xlsx" -Append -Title "App registration secrets report" -TitleBold -TitleSize 14 -AutoSize -FreezeTopRow -MaxAutoSizeRows 12 
