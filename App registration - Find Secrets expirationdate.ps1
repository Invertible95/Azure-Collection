###In Development###
# Install the required module if not already installed
# Install-Module -Name Microsoft.Graph.Authentication -Force -AllowClobber
# Install-Module -Name ImportExcel -Force -Scope CurrentUser

# Import the required modules
Import-Module Microsoft.Graph.Authentication
Import-Module ImportExcel

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Application.ReadWrite.All"

$allApplications = @()
$nextLink = $null

do {
    $applicationsPage = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications\$($nextLink -replace '\?', '&')"
 
    $allApplications += $applicationsPage.Value
 
    $nextLink = $applicationsPage.'@odata.nextLink'
} while ($nextLink)

$retArr = @()

foreach ($Application in $allApplications) {
    

    $SecretURI = "https://graph.microsoft.com/v1.0/applications/$($Application.id)/passwordCredentials"
    try {
        $FindSecretDates = Invoke-MgGraphRequest -Method GET -Uri $SecretURI
    }
    catch {
        Write-Host "$_.ErrorDetails"
        continue
    }
        
    foreach ($expireDate in $FindSecretDates.Value) {
        
        $expiryDateTime = [datetime]$expireDate.endDateTime
        $expiryDate = $expiryDateTime.Date
        
        if ($expireDate -ne $null) {
            $daysUntilExpiry = ($expiryDate - (Get-Date).Date).Days
            if ($daysUntilExpiry -gt 30) {
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
$excelDate = (Get-Date -Format "yyyy-MM-dd")
$retArr | Export-Excel -Path "C:\temp\Report $excelDate.xlsx" -Append -Title "App registration secrets report" -TitleBold -TitleSize 14 -AutoSize -FreezeTopRow -MaxAutoSizeRows 12 
