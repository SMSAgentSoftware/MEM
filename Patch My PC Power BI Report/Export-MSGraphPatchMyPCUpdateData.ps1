###########################################################################
## Azure automation runbook PowerShell script to export app install data ##
## for Intune apps and updates created by Patch My PC and dump it to     ##
## Azure Blob storage where it can be used as a datasource for Power BI. ##
###########################################################################

## Module Requirements ##
# Az.Accounts
# Az.Storage
# MSAL.PS (if using Runas account)

# Set some variables
$ProgressPreference = 'SilentlyContinue'
$ResourceGroup = "<my-resource-group-name>" # Reource group that hosts the storage account
$StorageAccount = "<my-storage-account-name>" # Storage account name
$Container = "patchmypc-powerbi" # Container name
$script:Destination = "$env:Temp"
$script:UpdateDeviceInstallStatusResults = New-Object System.Collections.ArrayList
$script:AppDeviceInstallStatusResults = New-Object System.Collections.ArrayList
$script:UpdateExportRequests = New-Object System.Collections.ArrayList
$script:AppExportRequests = New-Object System.Collections.ArrayList


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

# Runas account
# Requires MSAL.PS module
<#
$connectionName = "AzureRunAsConnection"
$servicePrincipalConnection = Get-AutomationConnection -Name $connectionName 
$Cert = Get-Item Cert:\LocalMachine\Root\$($servicePrincipalConnection.CertificateThumbprint)
$script:accessToken = (Get-MsalToken -ClientID $servicePrincipalConnection.ApplicationId -ClientCertificate $Cert -TenantId $servicePrincipalConnection.TenantId -Scopes 'https://graph.microsoft.com/.default').AccessToken
#>

## Connect to Azure AD 
# Mmanaged Identity
$null = Connect-AzAccount -Identity

# Runas account
#$null = Connect-AzAccount -ServicePrincipal -Tenant $servicePrincipalConnection.TenantId -ApplicationId $servicePrincipalConnection.ApplicationId -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint 



###############
## FUNCTIONS ##
###############
# Function to export a device install status report
Function Export-StatusReport {
    Param($ReportOutputName,$ReportEntityName,$ApplicationData)

    # Some variables
    $reporturl = "deviceManagement/reports/$ReportEntityName"
    $headers = @{
        "Content-Type" = "application/json"
    }
    $DataTable = [System.Data.DataTable]::new()

    # Prepare the apps in batches of 20 due to the current limitation of batching with Graph
    [int]$SkipValue = 0
    $BatchArray = [System.Collections.ArrayList]::new()
    do {
        $batch = $ApplicationData | Select -First 20 -Skip $SkipValue
        [void]$BatchArray.Add($batch)
        $SkipValue = $SkipValue + 20
    } until ($SkipValue -ge $ApplicationData.Count)

    # Process each batch
    foreach ($batch in $BatchArray)
    {
        $requests = @()
        [int]$Id = 1

        # generate a request for each app in the batch
        foreach ($app in $batch)
        {
            $body = @{
                filter = "(ApplicationId eq '$($App.id)')"
                top = 3000
            }
            $requesthash = [ordered]@{
                id = $Id.ToString()
                method = "POST"
                url = $reporturl
                body = $Body
                headers = $headers
            }
            $requests += $requesthash
            $Id ++
        }

        # Convert the requests to JSON
        $requestbase = @{
            requests = $requests
        }
        $JsonBase = $requestbase | ConvertTo-Json -Depth 3

        # Send the batch
        $URL = "https://graph.microsoft.com/beta/`$batch"
        $batchheaders = @{'Authorization'="Bearer " + $accessToken; 'Accept'="application/json"}
        $WebRequest = Invoke-WebRequest -Uri $URL -Method POST -Headers $batchheaders -Body $JsonBase -ContentType "application/json" -UseBasicParsing
        $responses = ($WebRequest.Content | ConvertFrom-Json).responses | Sort-Object -Property id

        # process the responses into a datatable
        foreach ($response in $responses)
        {
            If ($response.status -eq 200)
            {
                $JSONresponse = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($response.body)) | ConvertFrom-Json
                If ($DataTable.Columns.Count -eq 0)
                {
                    foreach ($column in $JSONresponse.Schema)
                    {
                        [void]$DataTable.Columns.Add($Column.Column)
                    }
                }
                if ($JSONresponse.values.Count -ge 1)
                {
                    foreach ($value in $JSONresponse.Values)
                    {
                        [void]$DataTable.Rows.Add($value)
                    }
                }
            }
        }
    }

    # Export the data
    $DataTable | Export-Csv -Path "$Destination\$ReportOutputName.csv" -NoTypeInformation -Force
}

# Function to call the Graph REST API with error handling
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

# Function to get the list of Pmp Apps and Updates from Graph
Function script:Get-PmpAppsList {
    $URL = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?`$filter=startswith(notes, 'Pmp')&`$expand=assignments&`$select=id,displayName,description,publisher,createdDateTime,lastModifiedDateTime,notes"
    $headers = @{'Authorization'="Bearer " + $accessToken}
    $Result = [System.Collections.ArrayList]::new()
    $GraphRequest = Invoke-LocalGraphRequest -URL $URL -Headers $headers -Method GET
    If ($GraphRequest.StatusCode -ne 200)
    {
        Return $GraphRequest
    }
    $Content = $GraphRequest.Content | ConvertFrom-Json
    [void]$Result.Add($Content.value)
    
    # Page through the next links if there are any
    If ($Content.'@odata.nextLink')
    {
        Do {
            Write-Output "Processing $($Content.'@odata.nextLink')"
            $GraphRequest = Invoke-LocalGraphRequest -URL $Content.'@odata.nextLink' -Headers $headers -Method GET
            If ($GraphRequest.StatusCode -ne 200)
            {
                Return $GraphRequest
            }
            $Content = $GraphRequest.Content | ConvertFrom-Json
            [void]$Result.Add($Content.value)
        }
        While ($null -ne $Content.'@odata.nextLink')
    }
    Return $Result
}

# Function to export the Pmp Apps and Updates to CSV
Function Export-PmpAppsList {

    $Result = Get-PmpAppsList
    If ($Result.GetType().BaseType -eq [System.Net.WebResponse])
    {
        If ($Result.StatusCode.value__ -eq 504)
        {
            # Server timeout encountered, lets try again a couple of times
            Write-Warning -Message "Http 504 (gateway timeout) encountered while getting Pmp apps list. Retrying up to 3 times."
            [int]$RetryAttempts = 0
            do {
                $RetryAttempts ++ 
                Start-Sleep -Seconds 5
                $Result = Get-PmpAppsList 
            }
            until ($RetryAttempts -gt 3 -or $Result.GetType().BaseType -ne [System.Net.WebResponse])
        }
        If ($Result.GetType().BaseType -eq [System.Net.WebResponse])
        {
            throw "Http error encountered from Graph API. Status code: $($Result.StatusCode.value__). Status description: $($Result.StatusDescription)."
            Exit 1
        }
    }

    # Remove some unwanted properties
    $Results = $Result | Select -Property * -ExcludeProperty '@odata.type','assignments@odata.context'
    
    # Add the customised results to a datatable
    $DataTable = [System.Data.DataTable]::new()
    foreach ($column in ($Results | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name))
    {
        [void]$DataTable.Columns.Add($Column)
    }
    foreach ($Result in $Results)
    {
        [void]$DataTable.Rows.Add(
            $Result.assignments.count,  
            $Result.createdDateTime,  
            $Result.description,
            $Result.displayName,
            $Result.Id,
            $Result.lastModifiedDateTime,
            $Result.notes,
            $Result.publisher
        )
    }
    
    # Separate Intune apps and Intune updates and export only those with assignments
    [array]$script:PmpApps = $DataTable.Select("notes LIKE 'PmpAppId:*' AND assignments >= 1") 
    [array]$script:PmpUpdates = $DataTable.Select("notes LIKE 'PmpUpdateId:*' AND assignments >= 1") 
    $PmpApps | Export-Csv -Path $Destination\PmpApps.csv -NoTypeInformation -Force
    $PmpUpdates | Export-Csv -Path $Destination\PmpUpdates.csv -NoTypeInformation -Force
}


# Function to create an exportJob request in Graph
Function script:New-MSGraphExportJob {
    Param($ReportName,$Filter)

    $bodyHash = [ordered]@{
        reportName = $ReportName
        filter = $Filter
    }
    $bodyJson = $bodyHash | ConvertTo-Json -Depth 3

    $URL = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs"
    $Headers = @{'Authorization'="Bearer " + $accessToken; 'Accept'="application/json"}
    $GraphRequest = Invoke-LocalGraphRequest -URL $URL -Headers $Headers -Method POST -Body $bodyJson -ContentType "application/json"

    Return $GraphRequest
}

# Function to receive an export Job from a request
Function script:Receive-MSGraphExportJob {
    Param($ExportJobRequest)
    $Id = ($ExportJobRequest.Content | ConvertFrom-Json).Id
    $AppID = ($ExportJobRequest.Content | convertfrom-json).filter.split()[-1].Replace("'",'').Replace(")",'')
    $FileName = "$AppId.zip"
    do {
        $URL = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$Id')"
        $Headers = @{'Authorization'="Bearer " + $accessToken; 'Accept'="application/json"}
        $WebResponse = Invoke-LocalGraphRequest -URL $URL -Headers $Headers -Method GET
        $ReponseJson = $WebResponse.Content | ConvertFrom-Json
        $Status = $ReponseJson.status
        Start-Sleep -Seconds 1
    }
    Until ($Status -eq "Completed")

    $DownloadUrl = $ReponseJson.url   
    try {
        $DownloadRequest = Invoke-WebRequest -Uri $DownloadUrl -OutFile "$Destination\$FileName" -UseBasicParsing -PassThru
    }
    catch {
        $DownloadRequest = $_.Exception.Response
    }
    Return $DownloadRequest
}

# Function to filter out duplicates for an export job from Graph
Function script:Remove-MSGraphExportJobDuplicates {
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
        $Others = $Array | Sort LastModifiedDateTime -Descending | Select -Skip 1
        foreach ($item in $Others)
        {
            $Collection.Remove($item)
        }
    }

    Return $Collection

}

# Function to request the export jobs
function Request-ExportJobs {
    param($AppIDs,$Type)
    foreach ($AppId in $AppIDs)
    {
        Write-host "Requesting $AppId"
        $Report = New-MSGraphExportJob -ReportName DeviceInstallStatusByApp -Filter "(ApplicationId eq '$AppId')"  
        switch ($Type)
        {
            "Apps" {[void]$AppExportRequests.Add($Report)}
            "Updates" {[void]$UpdateExportRequests.Add($Report)}
        }
    }
}

# Function to receive and export the export jobs
Function Receive-ExportJobs{
    param($Type)
    switch ($Type)
    {
        "Updates" {$ExportRequests = $UpdateExportRequests}
        "Apps" {$ExportRequests = $AppExportRequests}
    }
    foreach ($ExportRequest in $ExportRequests)
    {
        $Job = Receive-MSGraphExportJob $ExportRequest
        $AppID = ($ExportRequest.Content | convertfrom-json).filter.split()[-1].Replace("'",'').Replace(")",'')
        Write-host "Received $AppId"
        $FileName = "$AppId.zip"
        Unblock-File -Path "$Destination\$FileName"
        $CsvFile = (Expand-Archive -Path "$Destination\$FileName" -DestinationPath $Destination -Force -Verbose) 4>&1
        $CsvFileName = $CsvFile[-1].ToString().Split('\')[-1].Replace("'.",'')
        $File = Get-Childitem -Path $Destination\$CsvFileName -File
        [Array]$Results = Import-Csv $File.FullName -UseCulture    
        If (($Results.DeviceId | Select -Unique).Count -lt $Results.Count)
        {
            [System.Collections.ArrayList]$ArrayList = $Results
            [array]$UniqueResults = Remove-MSGraphExportJobDuplicates -Collection $ArrayList
            switch ($Type) {
                "Updates" {foreach ($UniqueResult in $UniqueResults){[void]$UpdateDeviceInstallStatusResults.Add($UniqueResult)}}
                "Apps" {foreach ($UniqueResult in $UniqueResults){[void]$AppDeviceInstallStatusResults.Add($UniqueResult)}}
            }
        }
        else 
        {
            switch ($Type) {
                "Updates" {foreach ($Result in $Results){[void]$UpdateDeviceInstallStatusResults.Add($Result)}}
                "Apps" {foreach ($Result in $Results){[void]$AppDeviceInstallStatusResults.Add($Result)}}
            }
        }
    }
    switch ($Type) {
        "Updates" {$UpdateDeviceInstallStatusResults | Export-CSV -Path $Destination\PmpUpdatesDeviceInstallStatusReport.csv -NoTypeInformation -Force}
        "Apps" {$AppDeviceInstallStatusResults | Export-CSV -Path $Destination\PmpAppsDeviceInstallStatusReport.csv -NoTypeInformation -Force}
    }
}

###############################################
## Export list of PMP applications in Intune ##
###############################################
Export-PmpAppsList



#################################
## Export the Overview reports ##
#################################
Export-StatusReport -ReportOutputName "PmpAppsStatusOverviewReport" -ReportEntityName "getAppStatusOverviewReport" -ApplicationData $PmpApps
Export-StatusReport -ReportOutputName "PmpUpdatesStatusOverviewStatusReport" -ReportEntityName "getAppStatusOverviewReport" -ApplicationData $PmpUpdates



############################################
## Request and Export the Install reports ##
############################################
Request-ExportJobs -AppIDs $PmpUpdates.Id -Type "Updates"
Request-ExportJobs -AppIDs $PmpApps.Id -Type "Apps"
Receive-ExportJobs -Type "Updates"
Receive-ExportJobs -Type "Apps"



##########################################
## UPLOAD REPORTS TO AZURE BLOB STORAGE ##
##########################################
$StorageAccount = Get-AzStorageAccount -Name $StorageAccount -ResourceGroupName $ResourceGroup
"PmpApps.csv","PmpUpdates.csv","PmpAppsDeviceInstallStatusReport.csv","PmpUpdatesDeviceInstallStatusReport.csv","PmpAppsStatusOverviewReport.csv","PmpUpdatesStatusOverviewStatusReport.csv" | foreach {
    try {
        $File = $_
        Write-Verbose "Uploading $File to Azure storage container $Container"
        $null = Set-AzStorageBlobContent -File "$env:temp\$File" -Container $Container -Blob $File -Context $StorageAccount.Context -Force -ErrorAction Stop
    }
    catch {
        Write-Error -Exception $_ -Message "Failed to upload $file to blob storage"
    } 
}