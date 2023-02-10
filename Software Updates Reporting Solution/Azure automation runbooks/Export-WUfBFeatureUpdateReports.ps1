##############################################################################
## Exports Intune Feature Update Reports for every feature update profile,  ##
## combines the data and uploads to a log analytics workspace               ##
##############################################################################

# Set variables
$WorkspaceID = "<workspaceId>" # WorkspaceID of the Log Analytics workspace
$PrimaryKey = "<primaryKey>" # Primary key of the Log Analytics workspace
$ProgressPreference = 'SilentlyContinue'

# Authenticate with MS Graph
# Permisions required:
    # For MS Graph, one of:
        # DeviceManagementConfiguration.ReadWrite.All, DeviceManagementApps.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All
    # For Log Analytics workspace:
        # Role assignment: Log Analytics Contributor
try 
{
    # Managed identity - for automation
    $null = Connect-AzAccount -Identity
    $script:GraphToken = (Get-AzAccessToken -ResourceTypeName MSGraph -ErrorAction Stop).Token

    # Or for interactive auth
    #$script:GraphToken = Connect-MSGraph -PassThru
}
catch 
{
    Write-Error "Failed to obtain access token! $($_.Exception.Message)"
    throw
}

#region Functions
# Function to call the Graph REST API with error handling
Function script:Invoke-WebRequestPro {
    Param ($URL,$Headers,$Method,$Body)
    If ($Method -eq "POST")
    {
        try 
        {
            $WebRequest = Invoke-WebRequest -Uri $URL -Method $Method -Headers $Headers -Body $Body -ContentType "application/json" -UseBasicParsing
        }
        catch 
        {
            $WebRequest = $_.Exception.Response
        }
    }
    else 
    {
        try 
        {
            $WebRequest = Invoke-WebRequest -Uri $URL -Method $Method -Headers $Headers -UseBasicParsing
        }
        catch 
        {
            $WebRequest = $_.Exception.Response
        } 
    }  
    Return $WebRequest
}

# Converts a byte array from an MS Graph response to a regular array
Function Convert-BytesResponseToArray {
    Param([byte[]]$ByteArray)
    $String = [System.Text.Encoding]::UTF8.GetString($ByteArray)
    $ConvertedJson = ConvertFrom-Json $String
    if ($ConvertedJson.TotalRowCount -ge 1)
    {
        $datatable = [System.Data.DataTable]::new()
        foreach ($Item in $ConvertedJson.Schema)
        {
            [void]$datatable.Columns.Add($Item.Column,$Item.PropertyType)
        }
        foreach ($Value in $ConvertedJson.Values)
        {
            $NewRow = $datatable.NewRow()
            $i = 0
            foreach ($Column in $datatable.Columns.ColumnName)
            {
                $NewRow["$Column"] = $Value[$i]
                $i ++
            }
            [void]$datatable.Rows.Add($NewRow)
        }
        return ($datatable | Select -Property @($datatable.Columns.ColumnName)) -as [array]
    }
    else 
    {
        return @()
    }
}

# Gets Intune FU policies
Function Get-IntuneFUReportFilters {
    param($Body)
    $URL = "https://graph.microsoft.com/beta/deviceManagement/reports/getReportFilters"
    $headers = @{'Authorization'="Bearer " + $GraphToken; 'Accept'="application/json"}
    $GraphRequest = Invoke-WebRequestPro -URL $URL -Headers $headers -Method POST -Body $Body
    return $GraphRequest 
}

# Creates a new Intune FU export job
Function New-IntuneFUExportJob {
    param($Body)
    $URL = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs"
    $headers = @{'Authorization'="Bearer " + $GraphToken}
    $GraphRequest = Invoke-WebRequestPro -URL $URL -Headers $headers -Method POST -Body $Body
    return $GraphRequest 
}

# Gets the status of an Intune FU export job
Function Get-IntuneFUExportJobStatus {
    param($id)
    $URL = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$id')"
    $headers = @{'Authorization'="Bearer " + $GraphToken}
    $GraphRequest = Invoke-WebRequestPro -URL $URL -Headers $headers -Method GET
    return $GraphRequest 
}

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
#endregion

################################
# Get feature update policies ##
################################
Write-Output "Retreiving feature update policies"
$Body = @{
    name = "FeatureUpdatePolicy"
} | ConvertTo-Json
$Response = Get-IntuneFUReportFilters -Body $Body
If ($Response.StatusCode -eq 429)
{
    $i = 0
    do {
        Write-Warning "Got a 429. Will try again in a minute."
        Start-Sleep -Seconds 60
        $Response = Get-IntuneFUReportFilters -Body $Body
        $i ++
    }
    Until ($Response.StatusCode -ne 429 -or $i -eq 5)
}
If ($i -eq 5)
{
    Write-Error "Received too many consecutive 429 responses."
    return
}
If ($Response.StatusCode -eq 200)
{
    $Bytes = $Response.Content
    If ($Bytes.Length -gt 1)
    {
        if ($Bytes -is [byte[]])
        {
            $ReportFilters = Convert-BytesResponseToArray -ByteArray $Bytes
        }
    }
}
else 
{
    Write-Error "Unexpected return code: $($Response.StatusCode)"    
    return
}

###############################
## Export the FU report data ##
###############################
$DeviceArray = [System.Collections.ArrayList]::new()

foreach ($ReportFilter in $ReportFilters)
{
    Write-Output "Requesting export job for policy '$($ReportFilter.PolicyName)'"
    $Body = @{
        reportName = "WindowsUpdatePerPolicyPerDeviceStatus"
        filter = "(PolicyId eq '$($ReportFilter.PolicyId)')"
        select = @("DeviceName","UPN","DeviceId","AADDeviceId","CurrentDeviceUpdateStatusEventDateTimeUTC","CurrentDeviceUpdateStatus","CurrentDeviceUpdateSubstatus","AggregateState","LatestAlertMessage","LastWUScanTimeUTC")                            
        format = "csv"
        localizationType = "replaceLocalizableValues"
    } | ConvertTo-Json
    $Response = New-IntuneFUExportJob -Body $Body
    If ($Response.StatusCode -eq 429)
    {
        $i = 0
        do {
            Write-Warning "Got a 429. Will try again in a minute."
            Start-Sleep -Seconds 60
            $Response = New-IntuneFUExportJob -Body $Body
            $i ++
        }
        Until ($Response.StatusCode -ne 429 -or $i -eq 5)
    }
    If ($i -eq 5)
    {
        Write-Error "Received too many consecutive 429 responses."
        return
    }
    If ($Response.StatusCode -eq 201)
    {
        $Content = $Response.Content | ConvertFrom-Json
        $JobId = $Content.id
        $x = 0

        # Poll for the export job status until complete
        do {
            $Response = Get-IntuneFUExportJobStatus -id $Jobid
            If ($Response.StatusCode -eq 429)
            {
                $i = 0
                do {
                    Write-Warning "Got a 429. Will try again in a minute."
                    Start-Sleep -Seconds 60
                    $Response = Get-IntuneFUExportJobStatus -Body $Body
                    $i ++
                }
                Until ($Response.StatusCode -ne 429 -or $i -eq 5)
            }
            If ($i -eq 5)
            {
                Write-Error "Received too many consecutive 429 responses."
                return
            }
            if ($Response.StatusCode -eq 200)
            {
                $Content = $Response.Content | ConvertFrom-Json
            }
            Start-Sleep -Seconds 10
        } until ($Content.status -eq "completed" -or $x -eq 60)
        
        If ($Content.status -eq "completed")
        {
            # Download the export file
            $DownloadURL = $Content.url
            $x = 0
            do {
                try 
                {                    
                    $RandomName = (New-Guid).ToString().Substring(0,8)
                    Invoke-WebRequest -UseBasicParsing -Uri $DownloadURL -OutFile "$env:Temp\$RandomName.zip"
                    $Success = $true
                }
                catch 
                {
                    Write-Error "Download failure: $($_.Exception.Message)"
                    $x ++
                    Start-Sleep -Seconds 2
                }
            } until ($Success -eq $true -or $x -eq 5)
            If ($x -eq 5)
            {
                Write-Error "Exceeded the maximum number of retries on the download"
                return
            }
            
            # Extract the csv from the zip, rename, load to memory
            If ([System.IO.File]::Exists("$env:Temp\$RandomName.zip"))
            {
                Expand-Archive -LiteralPath "$env:Temp\$RandomName.zip" -DestinationPath $env:Temp -Force
                Start-Sleep -Seconds 1
                Move-Item -Path "$env:temp\$($Content.id).csv" -Destination "$env:temp\$($ReportFilter.PolicyName).csv" -Force
                $Data = Import-Csv -LiteralPath "$env:temp\$($ReportFilter.PolicyName).csv" -UseCulture
                # Add the policy name and FU version
                foreach ($item in $Data)
                {
                    $item | Add-Member NoteProperty PolicyName -Value $ReportFilter.PolicyName
                    $item | Add-Member NoteProperty FeatureUpdateVersion -Value $ReportFilter.FeatureUpdateVersion
                }
                If ($Data -isnot [array])
                {
                    [void]$DeviceArray.Add($Data)
                }
                else 
                {
                    [void]$DeviceArray.AddRange($Data) 
                }                
            }
            else 
            {
                Write-Error "Downloaded report not found"    
                return
            }
        }
        else 
        {
            Write-Error "Timed out waiting for export job"    
            return
        }
    }
    else 
    {
        Write-Error "Unexpected return code: $($Response.StatusCode)"    
        return
    }
}

###############################
## Post data to LA workspace ##
###############################
Write-Output "Posting data to LA workspace"
# Add a date so we have a consistent timestamp on all entries for this post
$ExportDate = Get-Date ([DateTime]::UtcNow) -Format "s"
foreach ($entry in $DeviceArray)
{
    $entry | Add-Member -MemberType NoteProperty -Name ExportDate -Value $ExportDate
}
$Json = ConvertTo-Json $DeviceArray -Compress
Write-Output "Posting $($DeviceArray.count) entries at $(([math]::Round(([System.Text.Encoding]::UTF8.GetByteCount($Json) / 1MB),2))) MB"
$Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_IntuneFUStatus"
If ($Result.GetType().Name -eq "ErrorRecord")
{
    Write-Error -Exception $Result.Exception
}
else 
{
    $Result.StatusCode  
}
