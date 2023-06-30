##################################################################
## Azure Automation runbook to export Windows 11 readiness data ##
## from MS Graph (Intune) into Azure blob storage for use with  ##
## Power BI reporting                                           ##
##################################################################

##############
# Change log #
##############
# 2021-10-19  |  Removed the "managedBy" field from the results to enable easy elimination of duplicate results
# 2021-10-05  |  First release


## Module Requirements ##
# Az.Accounts
# Az.Storage


# Variables
$ResourceGroup = "<my-resource-group>" # Reource group that hosts the storage account
$StorageAccount = "<my-storage-account>" # Storage account name
$Container = "windows11readiness" # Container name
$ExportLocation = "$env:TEMP"
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'Continue'


# Graph web request function
Function Invoke-MyGraphGetRequest {
    Param ($URL)
    try {
            $WebRequest = Invoke-WebRequest -Uri $URL -Method GET -Headers $Headers -UseBasicParsing
    }
    catch {
        $WebRequest = $_.Exception.Response
    }
    Return $WebRequest
}


# Authenticate
$url = $env:IDENTITY_ENDPOINT  
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]" 
$headers.Add("X-IDENTITY-HEADER", $env:IDENTITY_HEADER) 
$headers.Add("Metadata", "True") 
$body = @{resource='https://graph.microsoft.com/' } 
$accessToken = (Invoke-RestMethod $url -Method 'POST' -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body $body ).access_token
$script:Headers = @{
    'Authorization' = "Bearer $accessToken"
}

$null = Connect-AzAccount -Identity


# Graph URI
$URI = "https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsWorkFromAnywhereMetrics('allDevices')/metricDevices?`$select=id%2cdeviceName%2cosDescription%2cosVersion%2cupgradeEligibility%2cazureAdJoinType%2cupgradeEligibility%2cramCheckFailed%2cstorageCheckFailed%2cprocessorCoreCountCheckFailed%2cprocessorSpeedCheckFailed%2ctpmCheckFailed%2csecureBootCheckFailed%2cprocessorFamilyCheckFailed%2cprocessor64BitCheckFailed%2cosCheckFailed&dtFilter=all"

# Get data from Graph with some error handling
$Response = Invoke-MyGraphGetRequest -URL $URI 
if ($Response.StatusCode -ne 200)
{
    Write-Warning "Graph request returned $($Response.StatusCode)). Retrying..."
    Start-Sleep -Seconds 30
    $RetryCount = 0
    do {
        $Response = Invoke-MyGraphGetRequest -URL $URI 
        If ($Response.StatusCode -ne 200) 
        {
            Write-Warning "Graph request returned $($Response.StatusCode)). Retrying..."
            $RetryCount ++
            Start-Sleep -Seconds 30
        }
    }
    Until ($Response.StatusCode -eq 200 -or $RetryCount -ge 10)
    If ($RetryCount -ge 10)
    {
        Write-Error "Gave up waiting for a success response to the Graph request."
        throw
    }
}

# Loop through the nextLinks until all data is retrieved
$JsonResponse = $Response.Content | ConvertFrom-Json
$DeviceData = $JsonResponse.value
If ($JsonResponse.'@odata.nextLink')
{
    do {
        $URI = $JsonResponse.'@odata.nextLink'
        $Response = Invoke-MyGraphGetRequest -URL $URI 
        if ($Response.StatusCode -ne 200)
        {
            Write-Warning "Graph request returned $($Response.StatusCode)). Retrying..."
            Start-Sleep -Seconds 60
            $RetryCount = 0
            do {
                $Response = Invoke-MyGraphGetRequest -URL $URI 
                If ($Response.StatusCode -ne 200) 
                {
                    Write-Warning "Graph request returned $($Response.StatusCode)). Retrying..."
                    $RetryCount ++
                    Start-Sleep -Seconds 60
                }
            }
            Until ($Response.StatusCode -eq 200 -or $RetryCount -ge 10)
            If ($RetryCount -ge 10)
            {
                Write-Error "Gave up waiting for a success response to the Graph request."
                throw
            }
        }
        $JsonResponse = $Response.Content | ConvertFrom-Json
        $DeviceData += $JsonResponse.value
    } until ($null -eq $JsonResponse.'@odata.nextLink')
}

# Create a unique result set of Windows 10 devices
$AllDevices = $DeviceData | where {$_.osVersion -like "10.0.*"}
$Properties = ($AllDevices[0] | Get-Member -MemberType NoteProperty).Name
$AllDevices = $AllDevices | Sort-Object -Property $Properties -Unique

# output some numbers
Write-Verbose "Total devices: $($AllDevices.count)"

# Export to CSV
$AllDevices | Export-CSV -Path "$ExportLocation\alldevices.csv" -NoTypeInformation -Force

# Upload to blob storage
$StorageAccount = Get-AzStorageAccount -Name $StorageAccount -ResourceGroupName $ResourceGroup
@(
    "alldevices.csv"    
) | foreach {
    try {
        $FileName = $_
        Write-Verbose "Uploading $FileName to Azure storage container $Container"
        $null = Set-AzStorageBlobContent -File "$ExportLocation\$FileName" -Container $Container -Blob $FileName -Context $StorageAccount.Context -Force -ErrorAction Stop
    }
    catch {
        Write-Error -Exception $_ -Message "Failed to upload $FileName to blob storage"
    } 
}
