##############################################################################################
## Azure automation runbook PowerShell script to export inventory data from                 ##
## Proactive remediations in Microsoft Intune and send it to an Azure Log Analytics         ##
## workspace where it can be used as a datasource for Power BI or an Azure Monitor workbook ##
##############################################################################################

## Module Requirements ##
# None

# Set some variables
$ProactiveRemediationsScriptGUID = "<GUID>" # GUID of the Proactive remediations script package
$WorkspaceID = Get-AutomationVariable -Name "WorkspaceID" # Saved as an encrypted variable in the automation account 
$PrimaryKey = Get-AutomationVariable -Name "PrimaryKey" # Saved as an encrypted variable in the automation account 
$LogName = "INV_SoftwareUpdates" # The name of the custom log in the LA workspace to send the data to

$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'Continue'
$TempFileName = "SoftwareUpdatesTemp"
$Destination = "$env:Temp"



####################
## AUTHENTICATION ##
####################
$url = $env:IDENTITY_ENDPOINT  
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]" 
$headers.Add("X-IDENTITY-HEADER", $env:IDENTITY_HEADER) 
$headers.Add("Metadata", "True") 
$body = @{resource='https://graph.microsoft.com/' } 
$script:accessToken = (Invoke-RestMethod $url -Method 'POST' -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body $body ).access_token
$script:authHeader = @{
    'Authorization' = "Bearer $accessToken"
}



###############
## FUNCTIONS ##
###############
# Function to make a web reqeust to Graph with error handling
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

# Function to get an export job from Graph
Function Get-MSGraphExportJob {
    Param($ReportName,$Filter,$FileName,$Destination)

    $bodyHash = [ordered]@{
        reportName = $ReportName
        filter = $Filter
    }
    $bodyJson = $bodyHash | ConvertTo-Json -Depth 3

    $URL = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs"
    $Headers = @{'Authorization'="Bearer " + $accessToken; 'Accept'="application/json"}
    $GraphRequest = Invoke-LocalGraphRequest -URL $URL -Headers $Headers -Method POST -Body $bodyJson -ContentType "application/json"
    If ($GraphRequest.StatusCode -ne 201)
    {
        Return $GraphRequest
    }

    $Id = ($GraphRequest.Content | ConvertFrom-Json).Id
    do {
        Start-Sleep -Seconds 5
        $URL = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$Id')"
        $Headers = @{'Authorization'="Bearer " + $accessToken; 'Accept'="application/json"}
        $WebResponse = Invoke-LocalGraphRequest -URL $URL -Headers $Headers -Method GET
        $ReponseJson = $WebResponse.Content | ConvertFrom-Json
        $Status = $ReponseJson.status
    }
    Until ($Status -eq "Completed")

    $DownloadUrl = $ReponseJson.url   
    try {
        $DownloadRequest = Invoke-WebRequest -Uri $DownloadUrl -OutFile "$Destination\$FileName.zip" -UseBasicParsing -PassThru
    }
    catch {
        $DownloadRequest = $_.Exception.Response
    }
    Return $DownloadRequest

}

# Function to remove duplicate entries per deviceId retaining the latest entry only, for an export job
Function Remove-MSGraphExportJobDuplicates {
    Param([System.Collections.ArrayList]$Collection)

    # Filter out duplicate DeviceIds into an array
    [array]$DeviceIDs = $Collection.DeviceId
    $NonDuplicatesHash = @{}
    $DuplicatesArray = [System.Collections.ArrayList]::new()
    foreach ($DeviceID in $DeviceIDs)
    {
        try {
            $NonDuplicatesHash.Add($DeviceID,0)
        }
        catch [System.Management.Automation.MethodInvocationException] {
            [void]$DuplicatesArray.Add($DeviceID)
        }
    }

    # Remove all but the latest (ModifiedTime) entry in the collection for each duplicate
    $DuplicatesArray = $DuplicatesArray | Select -Unique	
    foreach ($Duplicate in $DuplicatesArray)
    {
        $Array = $Collection.Where({$_.DeviceId -eq $Duplicate})
        #$Latest = $Array | Sort ModifiedTime -Descending | Select -First 1
        $Others = $Array | Sort ModifiedTime -Descending | Select -Skip 1
        foreach ($item in $Others)
        {
            $Collection.Remove($item)
        }
    }

    Return $Collection

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


#################
## MAIN SCRIPT ##
#################

# Get an export job from Graph from Proactive Remediations (DO data) and handle errors
$Report = Get-MSGraphExportJob -ReportName DeviceRunStatesByProactiveRemediation -Filter "PolicyId eq '$ProactiveRemediationsScriptGUID'" -FileName $TempFileName -Destination $Destination
If ($Report.GetType().BaseType -eq [System.Net.WebResponse] -or $Report.GetType().Name -eq "WebResponseObject")
{
    If ($Report.GetType().Name -eq "WebResponseObject")
    {
        If ($Report.StatusCode -eq 504)
        {
            # Server timeout encountered, lets try again a couple of times
            Write-Warning -Message "Http 504 (gateway timeout) encountered while getting Graph data. Retrying up to 3 times."
            [int]$RetryAttempts = 0
            do {
                $RetryAttempts ++ 
                Start-Sleep -Seconds 5
                $Report = Get-MSGraphExportJob -ReportName DeviceRunStatesByProactiveRemediation -Filter "PolicyId eq '$ProactiveRemediationsScriptGUID'" -FileName $FileName -Destination $Destination
            }
            until ($RetryAttempts -gt 3 -or $Report.StatusCode -eq 200)
        }
        ElseIf ($Report.StatusCode -ne 200)
        {
            throw "Http error encountered from Graph API. Status code: $($Report.StatusCode). Status description: $($Report.StatusDescription)."
            Exit 1
        }
    }
    else 
    {
        If ($Report.StatusCode.value__ -eq 504)
        {
            # Server timeout encountered, lets try again a couple of times
            Write-Warning -Message "Http 504 (gateway timeout) encountered while getting Graph data. Retrying up to 3 times."
            [int]$RetryAttempts = 0
            do {
                $RetryAttempts ++ 
                Start-Sleep -Seconds 5
                $Report = Get-MSGraphExportJob -ReportName DeviceRunStatesByProactiveRemediation -Filter "PolicyId eq '$ProactiveRemediationsScriptGUID'" -FileName $FileName -Destination $Destination
            }
            until ($RetryAttempts -gt 3 -or $Report.GetType().BaseType -ne [System.Net.WebResponse])
        }
        If ($Report.GetType().BaseType -eq [System.Net.WebResponse])
        {
            throw "Http error encountered from Graph API. Status code: $($Report.StatusCode.value__). Status description: $($Report.StatusDescription)."
            Exit 1
        }
    }
}

# Extract the CSV file from the exportJob and import it
Start-Sleep -Seconds 5
Unblock-File -Path "$Destination\$TempFileName.zip"
$CsvFile = (Expand-Archive -Path "$Destination\$TempFileName.zip" -DestinationPath $Destination -Force -Verbose) 4>&1
$CsvFileName = $CsvFile[-1].ToString().Split('\')[-1].Replace("'.",'')
$File = Get-Childitem -Path $Destination\$CsvFileName -File
Rename-Item -Path $File.FullName -NewName "$TempFileName.csv" -Force
[array]$ImportedResults = Import-Csv $Destination\$TempFileName.csv -UseCulture
[System.Collections.ArrayList]$ImportedResults = $ImportedResults | 
    where {$_.PreRemediationDetectionScriptOutput -ne "" -and $_.DetectionStatus -eq 3 } | 
    Select -Property DeviceId,ModifiedTime,PreRemediationDetectionScriptOutput,DeviceName

# Filter out duplicate entries keep only the most recent per device
# !! This can take some time and processing power on a large data set !!
$FilteredResults = Remove-MSGraphExportJobDuplicates -Collection $ImportedResults | Sort DeviceName

# Create a datatable to hold the results
# Add the column names for the data you are collecting from PR
$DataTable = [System.Data.DataTable]::new()
"DeviceId","DeviceName","SU_ScheduledRebootTime","SU_RebootRequired","SU_EngageReminderLastShownTime","SU_PendingRebootStartTime" | foreach {
    [void]$DataTable.Columns.Add($_)
}
Foreach ($Item in $FilteredResults)
{
    $Data = $Item.PreRemediationDetectionScriptOutput | ConvertFrom-json
    $SU_ScheduledRebootTime = $Data.SU_ScheduledRebootTime
    $SU_RebootRequired = $Data.SU_RebootRequired
    $SU_EngageReminderLastShownTime = $Data.SU_EngageReminderLastShownTime
    $SU_PendingRebootStartTime = $Data.SU_PendingRebootStartTime
    [void]$DataTable.Rows.Add(
        $Item.DeviceId,
        $Item.DeviceName,
        $SU_ScheduledRebootTime,
        $SU_RebootRequired,
        $SU_EngageReminderLastShownTime,
        $SU_PendingRebootStartTime
    )
}

$DataTable.DefaultView.Sort = "DeviceName"
$DataTable = $DataTable.DefaultView.ToTable($true)


# Send the data to the LA workspace
$Json = $DataTable | Select -Property @($DataTable.Columns.ColumnName) | ConvertTo-Json
Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($json)) -logType $LogName 