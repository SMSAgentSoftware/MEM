################################################################################################
## Azure automation runbook to export Secure Boot Inventory data from Intune to Azure storage ##
################################################################################################

## REQUIREMENTS ##
# Requires the MS provided Secure Boot inventory script deployed as an Intune remediation
# - https://support.microsoft.com/en-gb/topic/monitoring-secure-boot-certificate-status-with-microsoft-intune-remediations-6696a27b-fa09-4570-b112-124965adc87f
# Requires a managed identity on the Azure automation runbook
# Managed identity requires:
# - Storage Blob Data Contributor role on the Azure storage account
# - DeviceManagementConfiguration.ReadWrite.All / DeviceManagementApps.ReadWrite.All / DeviceManagementManagedDevices.ReadWrite.All application permissions in Microsoft Graph
# ref: https://docs.smsagent.blog/microsoft-endpoint-manager-reporting/automating-data-exports-from-microsoft-graph

#region ----------------------------------- Parameters -----------------------------------------
# The guid of your Intune remediation script
$IntuneRemediationsScriptGUID = "bbc275a3-2e41-1234-5678-5efa056ea293"
# Resource group that hosts the storage account
$ResourceGroup = "<MyResourceGroup>" 
# Storage account name
$StorageAccount = "<MyStorageAccount>" 
# Container name
$Container = "<MyContainerName>" 
# Temp working directory
$Destination = "$env:temp"
# File names
$TempFileName = "SecureBootInventoryTemp"
$FinalFileName = "SecureBootInventory" # Don't change if using the Power BI report
#Speeds up web requests
$ProgressPreference = 'SilentlyContinue'
#endregion -------------------------------------------------------------------------------------


#region ----------------------------------- Functions ------------------------------------------
# Simple function to make a web request with error handling
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

# Function to create & download an export job from MS Graph
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
Function script:Remove-MSGraphExportJobDuplicates {
    Param([System.Collections.ArrayList]$Collection)
    $GroupedCollection = $Collection | Group-Object -Property DeviceId
    foreach ($Item in ($GroupedCollection | Where {$_.Count -gt 1}))
    {
        $Others = $Item.Group | Sort LastModifiedDateTime -Descending | Select -Skip 1
        foreach ($entry in $Others)
        {
            $Collection.Remove($entry)
        }
    }
    Return $Collection
}
#endregion -------------------------------------------------------------------------------------


#region ----------------------------------- Authentication -------------------------------------
Write-Output "Authenticating..."
try 
{
    $null = Connect-AzAccount -Identity -ErrorAction Stop
    $script:accessToken = (Get-AzAccessToken -ResourceTypeName MSGraph -AsSecureString -ErrorAction Stop).Token | ConvertFrom-SecureString -AsPlainText
}
catch 
{
    throw "Unable to authenticate: $_"
}
#endregion -------------------------------------------------------------------------------------


#region ----------------------------------- Get export job -------------------------------------
# Get an export job from Graph from Intune Remediations and handle errors
Write-Output "Requesting an export job from MS Graph"
$Report = Get-MSGraphExportJob -ReportName DeviceRunStatesByProactiveRemediation -Filter "PolicyId eq '$IntuneRemediationsScriptGUID'" -FileName $TempFileName -Destination $Destination
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
                $Report = Get-MSGraphExportJob -ReportName DeviceRunStatesByProactiveRemediation -Filter "PolicyId eq '$IntuneRemediationsScriptGUID'" -FileName $FileName -Destination $Destination
            }
            until ($RetryAttempts -gt 3 -or $Report.StatusCode -eq 200)
        }
        ElseIf ($Report.StatusCode -ne 200)
        {
            throw "Http error encountered from Graph API. Status code: $($Report.StatusCode). Status description: $($Report.StatusDescription)."
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
                $Report = Get-MSGraphExportJob -ReportName DeviceRunStatesByProactiveRemediation -Filter "PolicyId eq '$IntuneRemediationsScriptGUID'" -FileName $FileName -Destination $Destination
            }
            until ($RetryAttempts -gt 3 -or $Report.GetType().BaseType -ne [System.Net.WebResponse])
        }
        If ($Report.GetType().BaseType -eq [System.Net.WebResponse])
        {
            throw "Http error encountered from Graph API. Status code: $($Report.StatusCode.value__). Status description: $($Report.StatusDescription)."
        }
    }
}
#endregion -------------------------------------------------------------------------------------


#region ----------------------------------- Extract data ---------------------------------------
Start-Sleep -Seconds 5
$t = Test-Path "$Destination\$TempFileName.zip"
if ($t -ne $true)
{
    throw "Exported zip file not found!"
}
Write-Output "Extracting data from download"
try 
{
    Unblock-File -Path "$Destination\$TempFileName.zip" -ErrorAction Stop
}
catch 
{
    throw "Failed to unblock file: $_"
}
try 
{
    Expand-Archive -Path "$Destination\$TempFileName.zip" -DestinationPath $Destination -Force -ErrorAction Stop
}
catch 
{
    throw "Failed to expand the zip archive: $_"
}
$File = Get-Childitem -Path $Destination -File -Filter *.csv -ErrorAction SilentlyContinue
if ($null -eq $file)
{
    throw "Exported csv file not found!"
}
Rename-Item -Path $File.FullName -NewName "$TempFileName.csv" -Force
[array]$ImportedResults = Import-Csv $Destination\$TempFileName.csv -UseCulture
[System.Collections.ArrayList]$ImportedResults = $ImportedResults | 
    where {$_.PreRemediationDetectionScriptOutput -ne "" -and $_.DetectionStatus -in 3, 4 } | 
    Select -Property DeviceId,ModifiedTime,PreRemediationDetectionScriptOutput,DeviceName
If ($ImportedResults.Count -ge 1)
{
    Write-Output "Extracted $($ImportedResults.Count) results"
}
else
{
    throw "No results found!"
}
#endregion -------------------------------------------------------------------------------------


#region ----------------------------------- Process data ---------------------------------------
# Filter out duplicate entries keep only the most recent per device
Write-Output "Deduping output"
$FilteredResults = Remove-MSGraphExportJobDuplicates -Collection $ImportedResults | Sort DeviceName
Write-Output "Filtered $($FilteredResults.Count) unique results"

# Create a datatable to hold the results
Write-Output "Preparing a data table"
$DataTable = [System.Data.DataTable]::new()
$firstResult = $FilteredResults | Select -First 1
$fields = ($firstResult.PreRemediationDetectionScriptOutput | ConvertFrom-Json).psobject.properties.name
foreach ($field in $fields)
{
    [void]$DataTable.Columns.Add($field)
}

Foreach ($Item in $FilteredResults)
{
    $Data = $Item.PreRemediationDetectionScriptOutput | ConvertFrom-json
    $newRow = $DataTable.NewRow()
    foreach ($field in $fields)
    {
        $newRow.$field = $Data.$field
    }
    $DataTable.Rows.Add($newRow)
}

$DataTable.DefaultView.Sort = "Hostname"
$DataTable = $DataTable.DefaultView.ToTable($true)
#endregion -------------------------------------------------------------------------------------


#region ----------------------------------- Create and upload csv ------------------------------
Write-Output "Exporting data table to csv"
$DataTable | Export-Csv -Path $Destination\$FinalFileName.csv -NoTypeInformation -Force

Write-Output "Uploading csv"
try 
{
    $StorageAccount = Get-AzStorageAccount -Name $StorageAccount -ResourceGroupName $ResourceGroup -ErrorAction Stop
    $null = Set-AzStorageBlobContent -File "$env:temp\$FinalFileName.csv" -Container $Container -Blob "$FinalFileName.csv" -Context $StorageAccount.Context -Force -ErrorAction Stop
    Write-Output 'Successfully uploaded csv to Azure storage account'
}
catch 
{
    throw "Failed to upload csv to Azure storage account: $_"
}
#endregion -------------------------------------------------------------------------------------