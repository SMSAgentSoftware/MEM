############################################################
# Example Azure Automation Runbook for exporting data from #
# MS Graph and sending it to a custom Azure Monitor log    #
############################################################


# Variables
$ProgressPreference = 'SilentlyContinue'
$WorkspaceID = Get-AutomationVariable -Name "WorkspaceID" # Saved as an encrypted variable in the automation account 
$PrimaryKey = Get-AutomationVariable -Name "PrimaryKey" # Saved as an encrypted variable in the automation account
$LogType = "Devices" # The name of the log file that will be created or appended to


#############
# FUNCTIONS #
#############

# function to invoke a web request to MS Graph with error handling
Function script:Invoke-LocalGraphRequest {
    Param ($URL,$Headers,$Method,$Body,$ContentType)
    try {
        If ($Method -eq "Post")
        {
            $WebRequest = Invoke-WebRequest -Uri $URL -Method $Method -Headers $Headers -Body $Body -ContentType $ContentType -UseBasicParsing
        }
        else 
        {
            $WebRequest = Invoke-WebRequest -Uri $URL -Method $Method -Headers $Headers -UseBasicParsing
        }     
    }
    catch {
        $WebRequest = $_.Exception.Response
    }
    Return $WebRequest
}

# function to get managed Windows device data from MS Graph
Function Get-DeviceData {
    $URL = "https://graph.microsoft.com/beta/deviceManagement/manageddevices?`$filter=startsWith(operatingSystem,'Windows')&`$select=deviceName,Id,lastSyncDateTime,managementAgent,managementState,osVersion,skuFamily,deviceEnrollmentType,emailAddress,model,manufacturer,serialNumber,userDisplayName,joinType"
    $headers = @{'Authorization'="Bearer " + $accessToken}
    $GraphRequest = Invoke-LocalGraphRequest -URL $URL -Headers $headers -Method GET
    If ($GraphRequest.StatusCode -ne 200)
    {
        Return $GraphRequest
    }
    $JsonResponse = $GraphRequest.Content | ConvertFrom-Json
    $DeviceData = $JsonResponse.value
    If ($JsonResponse.'@odata.nextLink')
    {
        do {
            $URL = $JsonResponse.'@odata.nextLink'
            $GraphRequest = Invoke-LocalGraphRequest -URL $URL -Headers $headers -Method GET
            If ($GraphRequest.StatusCode -ne 200)
            {
                Return $GraphRequest
            }
            $JsonResponse = $GraphRequest.Content | ConvertFrom-Json
            $DeviceData += $JsonResponse.value
        } until ($null -eq $JsonResponse.'@odata.nextLink')
    }
    Return $DeviceData
}

# Create the function to create the authorization signature
# ref https://docs.microsoft.com/en-us/azure/azure-monitor/logs/data-collector-api
Function Build-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource)
{
    $xHeaders = "x-ms-date:" + $date
    $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource

    $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
    $keyBytes = [Convert]::FromBase64String($sharedKey)

    $sha256 = New-Object System.Security.Cryptography.HMACSHA256
    $sha256.Key = $keyBytes
    $calculatedHash = $sha256.ComputeHash($bytesToHash)
    $encodedHash = [Convert]::ToBase64String($calculatedHash)
    $authorization = 'SharedKey {0}:{1}' -f $customerId,$encodedHash
    return $authorization
}

# Create the function to create and post the request
# ref https://docs.microsoft.com/en-us/azure/azure-monitor/logs/data-collector-api
Function Post-LogAnalyticsData($customerId, $sharedKey, $body, $logType)
{
    $method = "POST"
    $contentType = "application/json"
    $resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $contentLength = $body.Length
    $TimeStampField = ""
    $signature = Build-Signature `
        -customerId $customerId `
        -sharedKey $sharedKey `
        -date $rfc1123date `
        -contentLength $contentLength `
        -method $method `
        -contentType $contentType `
        -resource $resource
    $uri = "https://" + $customerId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"

    $headers = @{
        "Authorization" = $signature;
        "Log-Type" = $logType;
        "x-ms-date" = $rfc1123date;
        "time-generated-field" = $TimeStampField;
    }

    $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing
    return $response.StatusCode
}


####################
## AUTHENTICATION ##
####################

## Get MS Graph access token 
# Managed Identity
$url = $env:IDENTITY_ENDPOINT  
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]" 
$headers.Add("X-IDENTITY-HEADER", $env:IDENTITY_HEADER) 
$headers.Add("Metadata", "True") 
$body = @{resource='https://graph.microsoft.com/' } 
$script:accessToken = (Invoke-RestMethod $url -Method 'POST' -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body $body ).access_token


#########################
## THUNDERBIRDS ARE GO ##
#########################

$Devices = Get-DeviceData
$Json = $Devices | ConvertTo-Json
Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($json)) -logType $logType 