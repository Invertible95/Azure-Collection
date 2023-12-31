# Unused enterprise apps in Azure.

# Connection details
$tenantID = $env:TENANT_ID
$clientID = $env:CLIENT_ID_EAPPS
$clientSecretID = $env:CLIENT_SECRET_EAPPS
$tokenResponse = $null  # Initialize to null
$headers = @{}  # Initialize as an empty hashtable

# Authenticate to MS Graph API and get headers
function Connect-toGraph {

    $tokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $clientID
        Client_Secret = $clientSecretID
    }
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $tokenBody
    $global:headers = @{
        "Authorization" = "Bearer $($tokenResponse.access_token)"
        "Content-type"  = "application/json"
    }
    # Debug output
    Write-Host "Access Token: $($tokenResponse.access_token)"
    Write-Host "Headers: $($headers)"
}
Connect-toGraph
# Function to retrieve information about Enterprise Apps from Microsoft Graph API
function Get-EnterpriseApps {
    # Define the initial URL for fetching Enterprise Apps
    $URLGetApplications = "https://graph.microsoft.com/beta/servicePrincipals?" +
    "`$top=300&" + # Limit the number of results per page to 300
    "`$filter=servicePrincipalType eq 'Application' and (tags/any(tag: tag eq 'WindowsAzureActiveDirectoryIntegratedApp'))&" +
    "`$select=appid,id,servicePrincipalNames"

    # Initialize an array to store the pages of Enterprise Apps
    $appPages = @()
    # Retrieve the first page of Enterprise Apps
    $Applications = Invoke-RestMethod -Method GET -Uri $URLGetApplications -Headers $headers

    # Add the Enterprise Apps from the first page to the result array
    $appPages += $Applications.value

    # Check if there are more pages of Enterprise Apps (pagination)
    while ($Applications.'@odata.nextLink') {
        # Fetch the next page of Enterprise Apps
        $Applications = (Invoke-RestMethod -Method GET -Uri $Applications.'@odata.nextLink' -Headers $headers)
        
        # Add the Enterprise Apps from the next page to the result array
        $appPages += $Applications.value
    }

    # Return the accumulated list of Enterprise Apps
    return $appPages
}

# Function to retrieve the count of interactive sign-ins within a specified time frame
function Get-SignInsInteractive { 
    param (
        [System.Collections.ArrayList] $Applications,
        [int] $timeFrameInDays
    )

    $endDate = Get-Date
    $startDate = $endDate.AddDays(-$timeFrameInDays)
    $signInCount = @{}  # Create a hashtable to store the count for each application
    # Calculate the total number of iterations
    $totalIterations = $Applications.Count
    $currentIteration = 0
    
    foreach ($application in $Applications) {
        $currentIteration++
        # Calculate progress percentage
        $progressPercentage = [math]::Round(($currentIteration / $totalIterations) * 100)
        
        # Construct progress status
        $progressMessage = "Fetching Interactive Sign-Ins for Application $($application.appDisplayName) ($currentIteration/$totalIterations) - App ID: $($application.appId)"
        Write-Progress -Status "Progress" -PercentComplete $progressPercentage -Activity $progressMessage
        
        # Construct the URL to fetch the count of sign-ins for the current application
        $URLSignInsForAppInteractive = "https://graph.microsoft.com/v1.0/auditLogs/signIns?" +
        "`$filter=appId eq '$($application.appId)' " + # Filter by the application's ID
        "and createdDateTime ge $($startDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) " + # Start date filter
        "and createdDateTime le $($endDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))" +
        " and id eq '$($application.Id)'"  # End date filter

        try {
            # Invoke a REST API GET request to fetch the count of sign-ins for the current application
            $signIns = Invoke-RestMethod -Method GET -Uri $URLSignInsForAppInteractive -Headers $headers

            # Add the count of sign-ins to the result hashtable
            $signInCount[$application.appId] = @{
                count       = $signIns.value.Count
                DisplayName = $application.appDisplayName
                ID = $application.Id
            }
        }
        catch {
            # Handle any errors or null responses
            $signInCount[$application.appId] = @{
                Count       = 0
                DisplayName = $application.appDisplayName
            }
            Write-Host "Error fetching sign-ins for $($application.appId). Count set to 0." -ForegroundColor Red
        }

        # Introduce a delay to avoid rate limiting
        Start-Sleep -Seconds 1
    }
    Write-Progress -Status "Progress" -Completed
    return $signInCount
}

# Function to retrieve the count of non-interactive sign-ins within a specified time frame
function Get-SignInsNonInteractive {
    param (
        [System.Collections.ArrayList] $Applications,
        [int] $timeFrameInDays
    )

    $endDate = Get-Date
    $startDate = $endDate.AddDays(-$timeFrameInDays)
    $signInCount = @{}
    # Calculate the total number of iterations
    $totalIterations = $Applications.Count
    $currentIteration = 0
    
    foreach ($application in $Applications) {
        $currentIteration++
        # Calculate progress percentage
        $progressPercentage = [math]::Round(($currentIteration / $totalIterations) * 100)
        
        # Construct progress status
        $progressMessage = "Fetching Non-Interactive Sign-Ins for Application $($application.appDisplayName) ($currentIteration/$totalIterations) - App ID: $($application.appId)"
        Write-Progress -Status "Progress" -PercentComplete $progressPercentage -Activity $progressMessage
        
        # Construct the URL to fetch the count of sign-ins for the current application
        $URLSignInsForAppNonInteractive = "https://graph.microsoft.com/beta/auditLogs/signins?" +
        "&`$filter=(signInEventTypes/any(t: t eq 'nonInteractiveUser'))" +
        " and appId eq '$($application.appId)'" +
        " and createdDateTime ge $($startDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))" +
        " and createdDateTime le $($endDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))" +
        " and id eq '$($application.Id)'"

        # Invoke a REST API GET request to fetch the count of sign-ins for the current application
        $signIns = Invoke-RestMethod -Method Get -Uri $URLSignInsForAppNonInteractive -Headers $headers

        # Add the count of sign-ins to the result hashtable
        $signInCount[$application.appId] = @{
            count       = $signIns.value.Count
            DisplayName = $application.appDisplayName
            ID = $application.Id
        }
       
        # Introduce a delay to avoid rate limiting
        Start-Sleep -Seconds 1
    }
    # Clear the progress bar once the operation is complete
    Write-Progress -Status "Progress" -Completed
    return $signInCount
}

# Set the time frame for sign-in retrieval (e.g., 1 day)
$timeFrameInDays = 1

# Retrieve Enterprise Apps and associated Sign-Ins
$enterPriseApps = Get-EnterpriseApps
$interActiveSignIns = Get-SignInsInteractive -Applications $enterPriseApps -timeFrameInDays $timeFrameInDays
Connect-toGraph
$nonInteractiveSignIns = Get-SignInsNonInteractive -Applications $enterPriseApps -timeFrameInDays $timeFrameInDays

Connect-MgGraph
# Function to remove inactive Enterprise Apps
function Remove-InactiveEnterpriseApps {

    try {
        $RemoveableApps = $interactiveSignIns.GetEnumerator() | Where-Object { ( $_.Value.Count -eq '0' ) } | `
            ForEach-Object {
            [PSCustomObject]@{
                AppId       = $_.Key
                Count       = $_.Value.Count
                DisplayName = $_.Value.DisplayName
                Id = $_.Value.Id
            }
        }
        $RemoveApps = $RemoveableApps.Id
    }
    catch {
        Write-Error -Verbose
    }

    foreach ($ServicePrincipalId in $RemoveApps) {
        Remove-MgServicePrincipal -ServicePrincipalId $ServicePrincipalId
    }

    try {
        $RemoveableApps2 = $nonInteractiveSignIns.GetEnumerator() | Where-Object { ( $_.Value.Count -eq '0' ) } | `
            ForEach-Object {
            [PSCustomObject]@{
                AppId       = $_.Key
                Count       = $_.Value.Count
                DisplayName = $_.Value.DisplayName
                Id = $_.Value.id
            }
        }
        $RemoveApps2 = $RemoveableApps2.Id
    }
    catch {
        Write-Error -Verbose
    }

    foreach ($ServicePrincipalId in $RemoveApps2) {
        Remove-MgServicePrincipal -ServicePrincipalId $ServicePrincipalId
    }
}

# Function to export sign-in counts to CSV files
function Export-SignInsToCSV {
    # Export the signInCounts hashtable to CSV for interactive sign-ins
    $interactiveSignIns.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{
            AppId       = $_.Key
            Count       = $_.Value.Count
            DisplayName = $_.Value.DisplayName
            Id = $_.Value.id
        }
    } | Export-Csv -Path "C:\temp\SignInCountsInteractive.csv" -NoTypeInformation -Encoding utf8

    # Export the signInCounts hashtable to CSV for non-interactive sign-ins
    $nonInteractiveSignIns.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{
            AppId       = $_.Key
            Count       = $_.Value.Count
            DisplayName = $_.Value.DisplayName
            Id = $_.Value.id
        }
    } | Export-Csv -Path "C:\temp\SignInCountsNonInteractive.csv" -NoTypeInformation -Encoding utf8  
}
