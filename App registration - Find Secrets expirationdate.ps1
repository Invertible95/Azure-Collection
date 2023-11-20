###In Development###
# Install the required module if not already installed
# Install-Module -Name Microsoft.Graph.Authentication -Force -AllowClobber

# Import the module
Import-Module Microsoft.Graph.Authentication

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
    

    $SecretURI = "https://graph.microsoft.com/v1.0/applications/$($Application.Id)/passwordCredentials"
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
            if ($daysUntilExpiry -le 30) {
                $myObject = [PSCustomObject]@{
                    ApplicationName = $($Application.displayName)
                    ExpiryDate      = $($expiryDate.ToString("yyyy-MM-dd"))
                }
                #$myObject = New-Object -TypeName psobject
                #Add-Member -InputObject $myObject -MemberType NoteProperty -Name ApplicationName -Value $($Application.displayName)
                #Add-Member -InputObject $myObject -MemberType NoteProperty -Name ExpiryDate -Value $($expiryDate.ToString("yyyy-MM-dd"))
                $retArr += $myObject
            } 
        } 
    }

}   
$retArr #| Export-Excel -Path "C:\temp\certExpires.xlsx" -Append
