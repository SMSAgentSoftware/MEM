###############################################################################################
## Azure automation runbook that builds trend data into separate tables in the log analytics ##
## analytics workspace. This offloads the trend data processing from the Power BI report     ##
###############################################################################################

## IMPORTANT! This runbook should be scheduled at least 5-10 mins after the the summarizer runbook has completed to allow time for data ingestion

##################
## AUTHENTICATE ##
##################
#region Authentication
$ResourceGroupName = "<ResourceGroupName>" # Name of the resource group containing your log analytics workspace
$WorkspaceName = "<WorkspaceName>" # The log analytics workspace name
$WorkspaceID = "<WorkspaceID>" # The WorkspaceID of the Log Analytics workspace
$PrimaryKey = "<PrimaryKey>" # The primary key of the log analytics workspace

$ProgressPreference = 'SilentlyContinue'

# Mmanaged Identity
$null = Connect-AzAccount -Identity

# Connect to LA workspace
$Workspace = Get-AzOperationalInsightsWorkspace -ResourceGroupName $ResourceGroupName -Name $WorkspaceName -ErrorAction Stop 

# Make sure the thread culture is US for consistency of dates. Applies only to the single execution.
If ([System.Globalization.CultureInfo]::CurrentUICulture.Name -ne "en-US")
{
    [System.Globalization.CultureInfo]::CurrentUICulture = [System.Globalization.CultureInfo]::new("en-US")
}
If ([System.Globalization.CultureInfo]::CurrentCulture.Name -ne "en-US")
{
    [System.Globalization.CultureInfo]::CurrentCulture = [System.Globalization.CultureInfo]::new("en-US")
}
#endregion

###############
## FUNCTIONS ##
###############
#region Functions
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

    try {
        $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing
    }
    catch {
        $response = $_#.Exception.Response
    }
    
    return $response
}
#endregion Functions


###################################
## POST CU COMPLIANCE TREND DATA ##
###################################
#region PostCUTrend
Write-Output "Querying trend data for CU Compliance"
$Query = @"
let SummarizationDate = SU_ClientComplianceStatus_CL | summarize max(SummarizationTime_t);
SU_ClientComplianceStatus_CL
| where SummarizationTime_t in (SummarizationDate)
| where LastSyncTime_t between (ago(30d)..now())
| where DisplayVersion_s != "Dev"
| project 
    SummarizationTime_t,
    IntuneDeviceID_g,
    LatestRegularUpdateStatus=LatestRegularUpdateStatus_s,
    LatestRegularComplianceState=ComplianceStateRegular_s,
    LatestRegularUpdateName=LatestRegularUpdateName_s,
    Windows_Release=Windows_Release_s,
    OSVersion=OSVersion_s,
    OSVersionSupportStatus=OSVersionSupportStatus_s,
    LastSyncTime=LastSyncTime_t
| summarize 
    Count=dcount(IntuneDeviceID_g,4) by SummarizationTime_t, 
    LatestRegularComplianceState, 
    LatestRegularUpdateStatus,
    LatestRegularUpdateName,
    Windows_Release,
    OSVersion,
    OSVersionSupportStatus,
    LastSyncTime=bin(LastSyncTime,7d)
| order by LastSyncTime desc
"@
try
{
    $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -IncludeStatistics -ErrorAction Stop
}
catch 
{
    Write-Error "Invocation of the Log Analytics query failed: $($_.Exception.Message)"
    Write-Output "Let's try the LA query again..."
    try 
    {
        $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -IncludeStatistics -ErrorAction Stop
    }
    catch 
    {
        Write-Error "Invocation of the Log Analytics query failed again: $($_.Exception.Message)"
        throw
    }
}

If ($null -ne $Result.Error)
{
    Write-Error $Result.Error.Message
    throw $Result.Error.Details.InnerError.Message
}
else 
{
    $TableStats = $Result.Statistics.query | Select-String "TableRowCount" | ConvertFrom-Json
    $CPUTime = ($Result.Statistics.query | Select-String "totalCPU" | ConvertFrom-Json).cpu.totalCpu
    Write-Output "LA Query Stats"
    Write-Output "=============="
    Write-Output "CPU time (hh:mm:ss): $CPUTime"
    Write-Output "Row count: $($TableStats.tableRowCount)"
    Write-Output "Table size (MB): $($TableStats.tableSize / 1MB)" 
}

Write-Output "Posting to SU_CUComplianceTrendLatest table"
$iResults = $Result.Results
$TrendArray = [System.Linq.Enumerable]::ToArray($iResults)
$PostedTime = Get-Date ([DateTime]::UtcNow) -Format "s"
foreach ($item in $TrendArray)
{
    $item | Add-Member -MemberType NoteProperty -Name PostedTime -Value $PostedTime -Force
}
$Json = $TrendArray | ConvertTo-Json -Compress
$Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_CUComplianceTrendLatest"
If ($Result.GetType().Name -eq "ErrorRecord")
{
    Write-Error -Exception $Result.Exception
}
else 
{
    $Result.StatusCode  
}
#endregion


############################################
## POST CU COMPLIANCE EXTENDED TREND DATA ##
############################################
#region PostCUTrendExtended
Write-Output "Querying trend data for CU Compliance Extended"
$Query = @"
let SummarizationDate = SU_ClientComplianceStatus_CL | summarize max(SummarizationTime_t);
let ComplianceSummary = SU_ClientComplianceStatus_CL
| where SummarizationTime_t in (SummarizationDate)
| where LastSyncTime_t between (ago(30d)..now())
| where DisplayVersion_s != "Dev"
| project 
    IntuneDeviceID_g,
    SummarizationTime_t,
    LatestRegularComplianceState=ComplianceStateRegular_s,
    LatestRegularUpdateName=LatestRegularUpdateName_s,
    LatestPreviewComplianceState=ComplianceStatePreview_s,
    LatestPreviewUpdateName=LatestPreviewUpdateName_s,
    LatestOutofBandComplianceState=ComplianceStateOutofBand_s,
    LatestOutofBandUpdateName=LatestOutofBandUpdateName_s,
    LatestRegularLess1ComplianceState=ComplianceStateRegularLess1_s,
    LatestRegularUpdateLess1Name=LatestRegularUpdateLess1Name_s,
    LatestPreviewLess1ComplianceState=ComplianceStatePreviewLess1_s,
    LatestPreviewUpdateLess1Name=LatestPreviewUpdateLess1Name_s,
    LatestOutofBandLess1ComplianceState=ComplianceStateOutofBandLess1_s,
    LatestOutofBandUpdateLess1Name=LatestOutofBandUpdateLess1Name_s,
    LatestRegularLess2ComplianceState=ComplianceStateRegularLess2_s,
    LatestRegularUpdateLess2Name=LatestRegularUpdateLess2Name_s,
    LatestPreviewLess2ComplianceState=ComplianceStatePreviewLess2_s,
    LatestPreviewUpdateLess2Name=LatestPreviewUpdateLess2Name_s,
    LatestOutofBandLess2ComplianceState=ComplianceStateOutofBandLess2_s,
    LatestOutofBandUpdateLess2Name=LatestOutofBandUpdateLess2Name_s,
    Windows_Release=Windows_Release_s,
    OSVersion=OSVersion_s,
    OSVersionSupportStatus=OSVersionSupportStatus_s,
    LastSyncTime=LastSyncTime_t
| summarize 
    Count=dcount(IntuneDeviceID_g,4) by SummarizationTime_t, 
    LatestRegularComplianceState, 
    LatestRegularUpdateName,
    LatestPreviewComplianceState,
    LatestPreviewUpdateName,
    LatestOutofBandComplianceState,
    LatestOutofBandUpdateName,
    LatestRegularLess1ComplianceState,
    LatestRegularUpdateLess1Name,
    LatestPreviewLess1ComplianceState,
    LatestPreviewUpdateLess1Name,
    LatestOutofBandLess1ComplianceState,
    LatestOutofBandUpdateLess1Name,
    LatestRegularLess2ComplianceState,
    LatestRegularUpdateLess2Name,
    LatestPreviewLess2ComplianceState,
    LatestPreviewUpdateLess2Name,
    LatestOutofBandLess2ComplianceState,
    LatestOutofBandUpdateLess2Name,
    Windows_Release,
    OSVersion,
    OSVersionSupportStatus,
    LastSyncTime=bin(LastSyncTime,7d);
union 
(ComplianceSummary
| extend UpdateType="Latest Security 'B'"
| project SummarizationTime_t,UpdateName=LatestRegularUpdateName,ComplianceState=LatestRegularComplianceState,Count,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime
| summarize Count=sum(Count) by SummarizationTime_t,UpdateName,ComplianceState,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime),
(ComplianceSummary
| extend UpdateType="Latest Preview"
| project SummarizationTime_t,UpdateName=LatestPreviewUpdateName,ComplianceState=LatestPreviewComplianceState,Count,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime
| summarize Count=sum(Count) by SummarizationTime_t,UpdateName,ComplianceState,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime),
(ComplianceSummary
| extend UpdateType="Latest Out-of-Band"
| project SummarizationTime_t,UpdateName=LatestOutofBandUpdateName,ComplianceState=LatestOutofBandComplianceState,Count,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime
| summarize Count=sum(Count) by SummarizationTime_t,UpdateName,ComplianceState,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime),
(ComplianceSummary
| extend UpdateType="Previous Security 'B'"
| project SummarizationTime_t,UpdateName=LatestRegularUpdateLess1Name,ComplianceState=LatestRegularLess1ComplianceState,Count,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime
| summarize Count=sum(Count) by SummarizationTime_t,UpdateName,ComplianceState,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime),
(ComplianceSummary
| extend UpdateType="Previous Preview"
| project SummarizationTime_t,UpdateName=LatestPreviewUpdateLess1Name,ComplianceState=LatestPreviewLess1ComplianceState,Count,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime
| summarize Count=sum(Count) by SummarizationTime_t,UpdateName,ComplianceState,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime),
(ComplianceSummary
| extend UpdateType="Previous Out-of-Band"
| project SummarizationTime_t,UpdateName=LatestOutofBandUpdateLess1Name,ComplianceState=LatestOutofBandLess1ComplianceState,Count,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime
| summarize Count=sum(Count) by SummarizationTime_t,UpdateName,ComplianceState,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime),
(ComplianceSummary
| extend UpdateType="Previous +1 Security 'B'"
| project SummarizationTime_t,UpdateName=LatestRegularUpdateLess2Name,ComplianceState=LatestRegularLess2ComplianceState,Count,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime
| summarize Count=sum(Count) by SummarizationTime_t,UpdateName,ComplianceState,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime),
(ComplianceSummary
| extend UpdateType="Previous +1 Preview"
| project SummarizationTime_t,UpdateName=LatestPreviewUpdateLess2Name,ComplianceState=LatestPreviewLess2ComplianceState,Count,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime
| summarize Count=sum(Count) by SummarizationTime_t,UpdateName,ComplianceState,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime),
(ComplianceSummary
| extend UpdateType="Previous +1 Out-of-Band"
| project SummarizationTime_t,UpdateName=LatestOutofBandUpdateLess2Name,ComplianceState=LatestOutofBandLess2ComplianceState,Count,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime
| summarize Count=sum(Count) by SummarizationTime_t,UpdateName,ComplianceState,UpdateType,Windows_Release,OSVersion,OSVersionSupportStatus,LastSyncTime)
| where ComplianceState != "N/A"
| order by UpdateName,LastSyncTime,Count
"@
try
{
    $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -IncludeStatistics -ErrorAction Stop
}
catch 
{
    Write-Error "Invocation of the Log Analytics query failed: $($_.Exception.Message)"
    Write-Output "Let's try the LA query again..."
    try 
    {
        $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -IncludeStatistics -ErrorAction Stop
    }
    catch 
    {
        Write-Error "Invocation of the Log Analytics query failed again: $($_.Exception.Message)"
        throw
    }
}

If ($null -ne $Result.Error)
{
    Write-Error $Result.Error.Message
    throw $Result.Error.Details.InnerError.Message
}
else 
{
    $TableStats = $Result.Statistics.query | Select-String "TableRowCount" | ConvertFrom-Json
    $CPUTime = ($Result.Statistics.query | Select-String "totalCPU" | ConvertFrom-Json).cpu.totalCpu
    Write-Output "LA Query Stats"
    Write-Output "=============="
    Write-Output "CPU time (hh:mm:ss): $CPUTime"
    Write-Output "Row count: $($TableStats.tableRowCount)"
    Write-Output "Table size (MB): $($TableStats.tableSize / 1MB)" 
}

Write-Output "Posting to SU_CUComplianceTrendExtended table"
$iResults = $Result.Results
$TrendArray = [System.Linq.Enumerable]::ToArray($iResults)
$PostedTime = Get-Date ([DateTime]::UtcNow) -Format "s"
foreach ($item in $TrendArray)
{
    $item | Add-Member -MemberType NoteProperty -Name PostedTime -Value $PostedTime -Force
}
$Json = $TrendArray | ConvertTo-Json -Compress
$Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_CUComplianceTrendExtended"
If ($Result.GetType().Name -eq "ErrorRecord")
{
    Write-Error -Exception $Result.Exception
}
else 
{
    $Result.StatusCode  
}
#endregion


###################################
## POST FU COMPLIANCE TREND DATA ##
###################################
#region PostFUTrend
Write-Output "Querying trend data for FU Compliance"
$Query = @"
let SummarizationDate = SU_ClientComplianceStatus_CL | summarize max(SummarizationTime_t);
SU_ClientComplianceStatus_CL 
| where SummarizationTime_t in (SummarizationDate)
| where LastSyncTime_t between (ago(30d)..now())
| where DisplayVersion_s != "Dev"
| project 
    FriendlyOSName=FriendlyOSName_s,
    Windows_Release=Windows_Release_s,
    CurrentPatchLevel=CurrentPatchLevel_s,
    CurrentBuildNumber=CurrentBuildNumber_s,
    DisplayVersion=DisplayVersion_s,
    WindowsReleaseandVersion=strcat(Windows_Release_s,"" "",DisplayVersion_s),
    EditionID=EditionID_s,
    IntuneDeviceID=IntuneDeviceID_g,
    ComputerName=ComputerName_s,
    SummarizationTime=SummarizationTime_t,
    LastSyncTime=LastSyncTime_t
| summarize count() by SummarizationTime,FriendlyOSName,WindowsReleaseandVersion,Windows_Release,CurrentBuildNumber,DisplayVersion,CurrentPatchLevel,EditionID,LastSyncTime=bin(LastSyncTime,7d)
| order by SummarizationTime desc,CurrentPatchLevel desc
"@
try
{
    $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -IncludeStatistics -ErrorAction Stop
}
catch 
{
    Write-Error "Invocation of the Log Analytics query failed: $($_.Exception.Message)"
    Write-Output "Let's try the LA query again..."
    try 
    {
        $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -IncludeStatistics -ErrorAction Stop
    }
    catch 
    {
        Write-Error "Invocation of the Log Analytics query failed again: $($_.Exception.Message)"
        throw
    }
}

If ($null -ne $Result.Error)
{
    Write-Error $Result.Error.Message
    throw $Result.Error.Details.InnerError.Message
}
else 
{
    $TableStats = $Result.Statistics.query | Select-String "TableRowCount" | ConvertFrom-Json
    $CPUTime = ($Result.Statistics.query | Select-String "totalCPU" | ConvertFrom-Json).cpu.totalCpu
    Write-Output "LA Query Stats"
    Write-Output "=============="
    Write-Output "CPU time (hh:mm:ss): $CPUTime"
    Write-Output "Row count: $($TableStats.tableRowCount)"
    Write-Output "Table size (MB): $($TableStats.tableSize / 1MB)" 
}

Write-Output "Posting to SU_FUComplianceTrend table"
$iResults = $Result.Results
$TrendArray = [System.Linq.Enumerable]::ToArray($iResults)
$PostedTime = Get-Date ([DateTime]::UtcNow) -Format "s"
foreach ($item in $TrendArray)
{
    $item | Add-Member -MemberType NoteProperty -Name PostedTime -Value $PostedTime -Force
}
$Json = $TrendArray | ConvertTo-Json -Compress
$Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_FUComplianceTrend"
If ($Result.GetType().Name -eq "ErrorRecord")
{
    Write-Error -Exception $Result.Exception
}
else 
{
    $Result.StatusCode  
}
#endregion