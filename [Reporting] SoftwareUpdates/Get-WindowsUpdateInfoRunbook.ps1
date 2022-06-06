############################################################################
##  AZURE AUTOMATION RUNBOOK TO EXPORT WINDOWS UPDATE RELEASE INFORMATION ##
##  FROM MS DOCS TO A LOG ANALYTICS WORKSPACE                             ##
############################################################################

$WorkspaceID = Get-AutomationVariable -Name WorkspaceID
$PrimaryKey = Get-AutomationVariable -Name PrimaryKey
$script:Destination = "$env:TEMP"
$ProgressPreference = 'SilentlyContinue'


# Make sure the thread culture is US for consistency of dates. Applies only to the single execution.
If ([System.Globalization.CultureInfo]::CurrentUICulture.Name -ne "en-US")
{
    [System.Globalization.CultureInfo]::CurrentUICulture = [System.Globalization.CultureInfo]::new("en-US")
}
If ([System.Globalization.CultureInfo]::CurrentCulture.Name -ne "en-US")
{
    [System.Globalization.CultureInfo]::CurrentCulture = [System.Globalization.CultureInfo]::new("en-US")
}


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

# Function to output a datatable containing the support status of W10/11 versions
Function New-SupportTable {

    # Populate table columns
    If ($script:EditionsDatatable.Columns.Count -eq 0)
    {
        "Windows Release","Version","StartDate","EndDate","SupportPeriodInDays","InSupport","SupportDaysRemaining","EditionFamily" | foreach {
            If ($_ -eq "SupportPeriodInDays" -or $_ -eq "SupportDaysRemaining")
            {
                [void]$EditionsDatatable.Columns.Add($_,[int])
            }
            else 
            {
                [void]$EditionsDatatable.Columns.Add($_) 
            }          
        }
    }
    
    # Windows release info URLs
    $URLs = @(  
        "https://docs.microsoft.com/en-us/windows/release-health/windows11-release-information"
        "https://docs.microsoft.com/en-us/windows/release-health/release-information"
    )

    # Process each URL
    foreach ($URL in $URLs)
    {
        If ($URL -match "11")
        {
            $WindowsRelease = "Windows 11"
        }
        else 
        {
            $WindowsRelease = "Windows 10"
        }

        Switch ($WindowsRelease)
        {
            "Windows 10" {$Outputfile = "winreleaseinfo.html"}
            "Windows 11" {$Outputfile = "win11releaseinfo.html"}
        }
        
        Invoke-WebRequest -URI $URL -OutFile $Destination\$Outputfile -UseBasicParsing
        $htmlarray = Get-Content $Destination\$Outputfile -ReadCount 0

        $OSBuilds = $htmlarray | Select-String -SimpleMatch "(OS build "
        [array]$Versions = @()
        foreach ($item in $OSBuilds)
        {
            $Versions += $item.Line.Split()[1].Trim()
        }

        $EditionFamilies = @(
            'Home, Pro, Pro Education and Pro for Workstations'
            'Enterprise, Education and IoT Enterprise'
        )

        # Process each Windows version
        foreach ($Version in $Versions)
        {
            $Line = $htmlarray | Select-String -SimpleMatch "<td>$Version" | Where {$_ -notmatch "<tr>"}
            if ($Line)
            {       
                Switch ($WindowsRelease)
                {
                    "Windows 10" {$ServicingOption1 = "Semi-Annual Channel";$ServicingOption2 = "General Availability Channel"}
                    "Windows 11" {$ServicingOption1 = "General Availability Channel";$ServicingOption2 = "General Availability Channel"}
                }
                $LineNumber = $Line.LineNumber
                If ($htmlarray[$LineNumber] -match $ServicingOption1 -or $htmlarray[$LineNumber] -match $ServicingOption2)
                {
                    [string]$StartDate = ($htmlarray[($LineNumber + 1)].Replace('<td>','').Replace('</td>','').Trim())
                    [string]$EndDate1 = ($htmlarray[($LineNumber + 4)].Replace('<td>','').Replace('</td>','').Trim())
                    [string]$EndDate2 = ($htmlarray[($LineNumber + 5)].Replace('<td>','').Replace('</td>','').Trim())
                    
                    foreach ($family in $EditionFamilies)
                    {
                        if ($family -match "Pro")
                        {
                            [string]$EndDate = $EndDate1
                        }
                        else 
                        {
                            [string]$EndDate = $EndDate2    
                        }

                        $StartDateDT = [datetime]::ParseExact($StartDate, 'yyyy-MM-dd', $null) | Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
                        If ($EndDate -notmatch "End")
                        {
                            $EndDateDT = [datetime]::ParseExact($EndDate, 'yyyy-MM-dd', $null) | Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
                            $SupportDaysRemaining = ([datetime]$EndDateDT - (Get-Date)).Days
                            If ($SupportDaysRemaining -lt 0)
                            {
                                $SupportDaysRemaining = 0
                            }
                            $InSupport = $EndDateDT -gt (Get-Date)
                            $SupportPeriodInDays = ([datetime]$EndDateDT - [datetime]$StartDateDT).Days
                            $StartDateFinal = ($StartDateDT | Get-Date -Format "yyyy MMMM dd").ToString()
                            $EndDateFinal = ($EndDateDT | Get-Date -Format "yyyy MMMM dd").ToString()
                        }
                        else 
                        {
                            $SupportDaysRemaining = 0
                            $InSupport = "False"
                            $SupportPeriodInDays = "0"
                            $StartDateFinal = ($StartDate | Get-Date -Format "yyyy MMMM dd").ToString()
                            $EndDateFinal = $EndDate
                        }

                        [void]$EditionsDatatable.Rows.Add(
                            $WindowsRelease,
                            $Version,
                            $StartDateFinal,
                            $EndDateFinal,
                            $SupportPeriodInDays,
                            $InSupport,
                            $SupportDaysRemaining,
                            $family
                        )
                    }
                }
                else 
                {           
                    foreach ($family in $EditionFamilies)
                    { 
                        [void]$EditionsDatatable.Rows.Add(
                            $WindowsRelease,
                            $Version,
                            "End of service",
                            "End of service",
                            0,
                            "False",
                            0,
                            $family
                        )
                    }
                }
            }
            else 
            {
                foreach ($family in $EditionFamilies)
                { 
                    [void]$EditionsDatatable.Rows.Add(
                        $WindowsRelease,
                        $Version,
                        "End of service",
                        "End of service",
                        0,
                        "False",
                        0,
                        $family
                    )
                }
            }
            Remove-Variable StartDate -Force -ErrorAction SilentlyContinue
            Remove-Variable EndDate -Force -ErrorAction SilentlyContinue
        }
    }

    # Sort the table
    $EditionsDatatable.DefaultView.Sort = "[Windows Release] desc,Version desc,EditionFamily asc"
    $EditionsDatatable = $EditionsDatatable.DefaultView.ToTable($true)

}

Function New-UpdateHistoryTable {
    
    # Populate table columns
    If ($script:UpdateHistoryTable.Columns.Count -eq 0)
    {
        "Windows Release","ReleaseDate","KB","OSBuild","OSBaseBuild","OSRevisionNumber","OSVersion","Type" | foreach {
            If ($_ -eq "ReleaseDate")
            {
                [void]$UpdateHistoryTable.Columns.Add($_,[DateTime])
            }
            ElseIf ($_ -eq "OSBaseBuild" -or $_ -eq "OSRevisionNumber")
            {
                [void]$UpdateHistoryTable.Columns.Add($_,[int])
            }
            else 
            {
                [void]$UpdateHistoryTable.Columns.Add($_)
            }          
        }
    }

    $URLs = @(
        "https://aka.ms/WindowsUpdateHistory"
        "https://aka.ms/Windows11UpdateHistory"
    )

    # Process each URL
    foreach ($URL in $URLs)
    {
        If ($URL -match "11")
        {
            $WindowsRelease = "Windows 11"
        }
        else 
        {
            $WindowsRelease = "Windows 10"
        }

        Switch ($WindowsRelease)
        {
            "Windows 10" {$Outputfile = "winupdatehistoryinfo.html"}
            "Windows 11" {$Outputfile = "win11updatehistoryinfo.html"}
        }

        $Response = Invoke-WebRequest -Uri $URL -UseBasicParsing -ErrorAction Stop
        $Response.Content | Out-file $Destination\$Outputfile -Force
        $htmlarray = Get-Content -Path $Destination\$Outputfile -ReadCount 0 
        $OSbuildsarray = $htmlarray | Select-string -SimpleMatch "OS Build" 
        $KBarray = @()
        foreach ($OSbuild in $OSbuildsarray)
        {
            $KBarray += $OSbuild.Line.Split('>')[1].Replace('</a','').Replace('&#x2014;',' - ')
        }
        [array]$KBarray = $KBarray | Where {$_ -notmatch "Mobile" -and $_ -notmatch "15254."} | Select -Unique

        foreach ($item in $KBarray)
        {
            $Date = $item.Split('-').Trim()[0]
            $KB = $item.Split('-').Trim()[1].Split()[0]
            If ($KB.Length -lt 8)
            {
                $KB = "$($item.Split('-').Trim()[1].Split()[0])" + "$($item.Split('-').Trim()[1].Split()[1])"
            }
            [array]$BuildNumbers = $item.Split().Split(',').Replace(')','') | Where {$_ -match '[0-9]*\.[0-9]+'}
            $Type = $item.Split(')')[1].Trim()
            If ($Type -eq "" -or $null -eq $Type)
            {
                $PatchTuesday = Get-PatchTuesday -ReferenceDate $Date
                If (($Date | Get-Date) -eq $PatchTuesday)
                {
                    $Type = "Regular"
                }
                else 
                {
                    $Type = "Preview" # Could be out-of-band - how to detect?
                }
            }

            foreach ($BuildNumber in $BuildNumbers)
            {
                [void]$UpdateHistoryTable.Rows.Add(
                    $WindowsRelease,
                    $Date,
                    $KB,
                    $BuildNumber,
                    $BuildNumber.Split('.')[0],
                    $BuildNumber.Split('.')[1],
                    ($VersionBuildTable.Select("[Windows Release]='$WindowsRelease' and Build='$($BuildNumber.Split('.')[0])'")).Version,
                    $Type)
            }
        }
    }

    # Sort the table
    $UpdateHistoryTable.DefaultView.Sort = "[Windows Release] desc, OSBaseBuild desc, KB desc"
    $UpdateHistoryTable = $UpdateHistoryTable.DefaultView.ToTable($true)
}

# Function to output a datatable containing the latest updates for each W10/11 version
Function New-LatestUpdateTable {
    "Windows Release",
    "OSBaseBuild",
    "OSVersion",
    "LatestUpdate",
    "LatestUpdate_KB",
    "LatestUpdate_ReleaseDate",
    "LatestRegularUpdate",
    "LatestRegularUpdate_KB",
    "LatestRegularUpdate_ReleaseDate",
    "LatestPreviewUpdate",
    "LatestPreviewUpdate_KB",
    "LatestPreviewUpdate_ReleaseDate",
    "LatestOutofBandUpdate",
    "LatestOutofBandUpdate_KB",
    "LatestOutofBandUpdate_ReleaseDate",
    "LatestRegularUpdateLess1",
    "LatestRegularUpdateLess1_KB",
    "LatestRegularUpdateLess1_ReleaseDate",
    "LatestPreviewUpdateLess1",
    "LatestPreviewUpdateLess1_KB",
    "LatestPreviewUpdateLess1_ReleaseDate",
    "LatestOutofBandUpdateLess1",
    "LatestOutofBandUpdateLess1_KB",
    "LatestOutofBandUpdateLess1_ReleaseDate",
    "LatestRegularUpdateLess2",
    "LatestRegularUpdateLess2_KB",
    "LatestRegularUpdateLess2_ReleaseDate",
    "LatestPreviewUpdateLess2",
    "LatestPreviewUpdateLess2_KB",
    "LatestPreviewUpdateLess2_ReleaseDate",
    "LatestOutofBandUpdateLess2",
    "LatestOutofBandUpdateLess2_KB",
    "LatestOutofBandUpdateLess2_ReleaseDate",
    "LatestUpdateType" | foreach {
        If ($_ -eq "OSBaseBuild")
        {
            [void]$LatestUpdateTable.Columns.Add($_,[int])
        }
        else 
        {
            [void]$LatestUpdateTable.Columns.Add($_)
        }      
    }
    $WindowsReleases = @(
        "Windows 10"
        "Windows 11"
    )
    foreach ($WindowsRelease in $WindowsReleases)
    {
        [array]$BuildVersions = $UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease'").OSBaseBuild | Select -Unique | Sort -Descending
        foreach ($BuildVersion in $BuildVersions)
        {
            $OSVersion = ($VersionBuildTable.Select("[Windows Release]='$WindowsRelease' and Build='$BuildVersion'")).Version
            $LatestRegularUpdate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Regular"} | Sort ReleaseDate -Descending | Select -First 1).OSBuild
            $LatestUpdate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Sort ReleaseDate -Descending | Select -First 1).OSBuild
            $LatestPreviewUpdate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Preview"} | Sort ReleaseDate -Descending | Select -First 1).OSBuild
            $LatestOutofBandUpdate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Out-of-band"} | Sort ReleaseDate -Descending | Select -First 1).OSBuild
            $LatestRegularUpdate_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Regular"} | Sort ReleaseDate -Descending | Select -First 1).KB
            $LatestUpdate_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Sort ReleaseDate -Descending | Select -First 1).KB
            $LatestPreviewUpdate_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Preview"} | Sort ReleaseDate -Descending | Select -First 1).KB
            $LatestOutofBandUpdate_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Out-of-band"} | Sort ReleaseDate -Descending | Select -First 1).KB
            $LatestRegularUpdate_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Regular"} | Sort ReleaseDate -Descending | Select -First 1).ReleaseDate
            $LatestUpdate_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Sort ReleaseDate -Descending | Select -First 1).ReleaseDate
            $LatestPreviewUpdate_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Preview"} | Sort ReleaseDate -Descending | Select -First 1).ReleaseDate
            $LatestOutofBandUpdate_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Out-of-band"} | Sort ReleaseDate -Descending | Select -First 1).ReleaseDate

            $LatestRegularUpdateLess1 = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Regular"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 1).OSBuild
            $LatestPreviewUpdateLess1 = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Preview"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 1).OSBuild
            $LatestOutofBandUpdateLess1 = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Out-of-band"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 1).OSBuild
            $LatestRegularUpdateLess1_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Regular"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 1).KB
            $LatestPreviewUpdateLess1_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Preview"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 1).KB
            $LatestOutofBandUpdateLess1_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Out-of-band"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 1).KB
            $LatestRegularUpdateLess1_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Regular"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 1).ReleaseDate
            $LatestPreviewUpdateLess1_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Preview"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 1).ReleaseDate
            $LatestOutofBandUpdateLess1_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Out-of-band"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 1).ReleaseDate

            $LatestRegularUpdateLess2 = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Regular"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 2).OSBuild
            $LatestPreviewUpdateLess2 = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Preview"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 2).OSBuild
            $LatestOutofBandUpdateLess2 = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Out-of-band"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 2).OSBuild
            $LatestRegularUpdateLess2_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Regular"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 2).KB
            $LatestPreviewUpdateLess2_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Preview"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 2).KB
            $LatestOutofBandUpdateLess2_KB = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Out-of-band"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 2).KB
            $LatestRegularUpdateLess2_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Regular"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 2).ReleaseDate
            $LatestPreviewUpdateLess2_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Preview"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 2).ReleaseDate
            $LatestOutofBandUpdateLess2_ReleaseDate = ($UpdateHistoryTable.Select("[Windows Release]='$WindowsRelease' and OSBaseBuild='$BuildVersion'") | Where {$_.Type -eq "Out-of-band"} | Sort ReleaseDate -Descending | Select -First 1 -Skip 2).ReleaseDate

            If ($LatestUpdate -eq $LatestRegularUpdate)
            { $LatestUpdateType = "Regular" }
            If ($LatestUpdate -eq $LatestPreviewUpdate)
            { $LatestUpdateType = "Preview" }
            If ($LatestUpdate -eq $LatestOutofBandUpdate)
            { $LatestUpdateType = "Out-of-band" }

            [void]$LatestUpdateTable.Rows.Add(
                $WindowsRelease,
                $BuildVersion,
                $OSVersion,
                $LatestUpdate,
                $LatestUpdate_KB,
                $LatestUpdate_ReleaseDate,
                $LatestRegularUpdate,
                $LatestRegularUpdate_KB,
                $LatestRegularUpdate_ReleaseDate,
                $LatestPreviewUpdate,
                $LatestPreviewUpdate_KB,
                $LatestPreviewUpdate_ReleaseDate,
                $LatestOutofBandUpdate,
                $LatestOutofBandUpdate_KB,
                $LatestOutofBandUpdate_ReleaseDate,
                $LatestRegularUpdateLess1,
                $LatestRegularUpdateLess1_KB,
                $LatestRegularUpdateLess1_ReleaseDate,
                $LatestPreviewUpdateLess1,
                $LatestPreviewUpdateLess1_KB,
                $LatestPreviewUpdateLess1_ReleaseDate,
                $LatestOutofBandUpdateLess1,
                $LatestOutofBandUpdateLess1_KB,
                $LatestOutofBandUpdateLess1_ReleaseDate,
                $LatestRegularUpdateLess2,
                $LatestRegularUpdateLess2_KB,
                $LatestRegularUpdateLess2_ReleaseDate,
                $LatestPreviewUpdateLess2,
                $LatestPreviewUpdateLess2_KB,
                $LatestPreviewUpdateLess2_ReleaseDate,
                $LatestOutofBandUpdateLess2,
                $LatestOutofBandUpdateLess2_KB,
                $LatestOutofBandUpdateLess2_ReleaseDate,
                $LatestUpdateType
            )
        }
    }
}

# Function to output a datatable referencing OS builds with versions
Function New-OSVersionBuildTable {

    # Windows release info URLs
    $URLs = @(
        "https://docs.microsoft.com/en-us/windows/release-health/release-information"
        "https://docs.microsoft.com/en-us/windows/release-health/windows11-release-information"
    )

    # Process each Windows release
    foreach ($URL in $URLs)
    {
        Invoke-WebRequest -URI $URL -OutFile $Destination\winreleaseinfo.html -UseBasicParsing
        $htmlarray = Get-Content $Destination\winreleaseinfo.html -ReadCount 0
        $versions = $htmlarray | Select-String -SimpleMatch "(OS build "
    
        If ($VersionBuildTable.Columns.Count -eq 0)
        {
            [void]$VersionBuildTable.Columns.Add("Windows Release")
            [void]$VersionBuildTable.Columns.Add("Version")
            [void]$VersionBuildTable.Columns.Add("Build",[int])
        }

        If ($URL -match "11")
        {
            $WindowsRelease = "Windows 11"
        }
        else 
        {
            $WindowsRelease = "Windows 10"
        }
    
        foreach ($version in $versions) 
        {
            $line = ($version.Line.split('>') | where {$_ -match "OS Build"}).TrimEnd('</strong')
            $ReleaseCode = $line.Split()[1]
            $Buildnumber = $line.Split()[-1].TrimEnd(')')
            [void]$VersionBuildTable.Rows.Add($WindowsRelease,$ReleaseCode,$Buildnumber)
        }
    }

    # Sort the table
    $VersionBuildTable.DefaultView.Sort = "[Windows Release] desc,Version desc"
    $VersionBuildTable = $VersionBuildTable.DefaultView.ToTable($true)
}

# Function to extract documented Windows Update error codes from MS docs
Function Get-WUErrorCodes {
    # create a table
    $ErrorCodeTable = [System.Data.DataTable]::new()
    $ErrorCodeTable.Columns.AddRange(@("ErrorCode","Description","Message","Category"))

    ################################
    ## PROCESS THE FIRST WEB PAGE ##
    ################################
    # scrape the web page
    $ProgressPreference = 'SilentlyContinue'
    $URL = "https://docs.microsoft.com/en-us/windows/deployment/update/windows-update-error-reference"
    $tempFile = [System.IO.Path]::GetTempFileName() 
    Invoke-WebRequest -URI $URL -OutFile $tempFile -UseBasicParsing 
    $htmlarray = Get-Content $tempFile -ReadCount 0
    [System.IO.File]::Delete($tempFile)

    # get the headers and data cells
    $headers = $htmlarray | Select-String -SimpleMatch "<h2 " | Where {$_ -notmatch "Feedback"}
    $dataCells = $htmlarray | Select-String -SimpleMatch "<td>"

    # process each header
    $i = 1
    do {
        foreach ($header in $headers)
        {
            $lineNumber = $header.LineNumber
            $nextHeader = $headers[$i]
            If ($null -ne $nextHeader)
            {
                $nextHeaderLineNumber = $nextHeader.LineNumber
                $cells = $dataCells | Where {$_.LineNumber -gt $lineNumber -and $_.LineNumber -lt $nextHeaderLineNumber}
            }
            else 
            {
                $cells = $dataCells | Where {$_.LineNumber -gt $lineNumber}  
            }

            # process each cell
            $totalCells = $cells.Count
            $t = 0
            do {
                $Row = $ErrorCodeTable.NewRow()
                "ErrorCode","Message","Description" | foreach {
                    $Row["$_"] = "$($cells[$t].ToString().Replace('<code>','').Replace('</code>','').Split('>').Split('<')[2])"
                    $t ++
                }
                $Row["Category"] = "$($header.ToString().Split('>').Split('<')[2])"
                [void]$ErrorCodeTable.Rows.Add($Row)
            }
            until ($t -ge ($totalCells -1))
            $i ++
        }
    }
    until ($i -ge $headers.count)

    #################################
    ## PROCESS THE SECOND WEB PAGE ##
    #################################
    # scrape the web page
    $URL = "https://docs.microsoft.com/en-us/windows/deployment/update/windows-update-errors"
    $tempFile = [System.IO.Path]::GetTempFileName() 
    Invoke-WebRequest -URI $URL -OutFile $tempFile -UseBasicParsing
    $htmlarray = Get-Content $tempFile -ReadCount 0
    [System.IO.File]::Delete($tempFile)

    # get the headers and data cells
    $headers = $htmlarray | Select-String -SimpleMatch "<h2 " | Where {$_ -notmatch "Feedback"}
    $dataCells = $htmlarray | Select-String -SimpleMatch "<td>"

    # process each header
    $i = 1
    do {
        foreach ($header in $headers)
        {
            $lineNumber = $header.LineNumber
            $nextHeader = $headers[$i]
            If ($null -ne $nextHeader)
            {
                $nextHeaderLineNumber = $nextHeader.LineNumber
                $cells = $dataCells | Where {$_.LineNumber -gt $lineNumber -and $_.LineNumber -lt $nextHeaderLineNumber}
            }
            else 
            {
                $cells = $dataCells | Where {$_.LineNumber -gt $lineNumber}  
            }

            # process each cell
            $totalCells = $cells.Count
            $t = 0
            do {
                $WebErrorCode = $header.ToString().Split('>').Split('<')[2].Replace('or ','').Replace("â€¯",' ').Split()
                If ($WebErrorCode.GetType().BaseType.Name -eq "Array")
                {
                    foreach ($Code in $WebErrorCode)
                    {
                        $Row = $ErrorCodeTable.NewRow()
                        $Row["ErrorCode"] = $Code.Trim()
                        "Message","Description" | foreach {
                            $Row["$_"] = "$($cells[$t].ToString().Split('>').Split('<')[2])"
                            $t ++
                        }
                        $Row["Category"] = "Common"
                        [void]$ErrorCodeTable.Rows.Add($Row)
                        1..2 | foreach {$t --}
                    }
                    1..2 | foreach {$t ++}
                }
                else {
                    $Row = $ErrorCodeTable.NewRow()
                    $Row["ErrorCode"] = $ErrorCode
                    "Message","Description" | foreach {
                        $Row["$_"] = "$($cells[$t].ToString().Split('>').Split('<')[2])"
                        $t ++
                    }
                    $Row["Category"] = "Common"
                    [void]$ErrorCodeTable.Rows.Add($Row)
                }
                $t ++
            }
            until ($t -ge ($totalCells -1))
            $i ++
        }
    }
    until ($i -ge $headers.count)
    Return $ErrorCodeTable
}

# Function to get current month's Patch Tuesday
# Thanks to https://github.com/tsrob50/Get-PatchTuesday/blob/master/Get-PatchTuesday.ps1
Function script:Get-PatchTuesday {
    [CmdletBinding()]
    Param
    (
      [Parameter(position = 0)]
      [ValidateSet("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")]
      [String]$weekDay = 'Tuesday',
      [ValidateRange(0, 5)]
      [Parameter(position = 1)]
      [int]$findNthDay = 2,
      [Parameter(position = 2)]
      [datetime]$ReferenceDate = [datetime]::NOW
    )
    # Get the date and find the first day of the month
    # Find the first instance of the given weekday
    $todayM = $ReferenceDate.Month.ToString()
    $todayY = $ReferenceDate.Year.ToString()
    [datetime]$strtMonth = $todayM + '/1/' + $todayY
    while ($strtMonth.DayofWeek -ine $weekDay ) { $strtMonth = $StrtMonth.AddDays(1) }
    $firstWeekDay = $strtMonth
  
    # Identify and calculate the day offset
    if ($findNthDay -eq 1) {
      $dayOffset = 0
    }
    else {
      $dayOffset = ($findNthDay - 1) * 7
    }
    
    # Return date of the day/instance specified
    $patchTuesday = $firstWeekDay.AddDays($dayOffset) 
    return $patchTuesday
}
#endregion


#################################
## UPDATE WU ERROR CODES TABLE ##
#################################
#region UpdateWUErrorCodes
# This runs twice a month just to keep the data from ageing past the data retention period in the LA workspace
If ([DateTime]::UtcNow.Day -eq 7 -or [DateTime]::UtcNow.Day -eq 21)
{
    $WUErrorCodes = Get-WUErrorCodes
    $Table = [System.Data.DataTable]::new()
    ($WUErrorCodes[0] | Get-Member -MemberType Property).Name | foreach {
        [void]$Table.Columns.Add($_)
    }
    foreach ($row in $WUErrorCodes)
    {
        $Table.ImportRow($row)
    }
    # Post the JSON to LA workspace
    $Json = $Table.Rows | Select ErrorCode,Description,Message,Category | ConvertTo-Json -Compress
    $Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_WUErrorCodes"
    If ($Result.GetType().Name -eq "ErrorRecord")
    {
        Write-Error -Exception $Result.Exception
    }
    else 
    {
        $Result.StatusCode  
    } 
}
#endregion


#############################################
## CREATE WINDOWS UPDATES REFERENCE TABLES ##
#############################################
#region CreateReferenceTables

$script:EditionsDatatable = [System.Data.DataTable]::new()
$script:UpdateHistoryTable = [System.Data.DataTable]::new()
$script:LatestUpdateTable = [System.Data.DataTable]::new()
$script:VersionBuildTable = [System.Data.DataTable]::new()

# Build the OS info tables
New-OSVersionBuildTable
New-SupportTable
New-UpdateHistoryTable
New-LatestUpdateTable

# Post the tables off to LA workspace
$Json = $EditionsDatatable.Rows | Select $EditionsDatatable.Columns.ColumnName | ConvertTo-Json -Compress
$Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_OSSupportMatrix"
If ($Result.GetType().Name -eq "ErrorRecord")
{
    Write-Error -Exception $Result.Exception
}
else 
{
    $Result.StatusCode  
} 

$Json = $UpdateHistoryTable.Rows | Select $UpdateHistoryTable.Columns.ColumnName | ConvertTo-Json -Compress
$Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_OSUpdateHistory"
If ($Result.GetType().Name -eq "ErrorRecord")
{
    Write-Error -Exception $Result.Exception
}
else 
{
    $Result.StatusCode  
} 

$Json = $LatestUpdateTable.Rows | Select $LatestUpdateTable.Columns.ColumnName | ConvertTo-Json -Compress
$Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_OSLatestUpdates"
If ($Result.GetType().Name -eq "ErrorRecord")
{
    Write-Error -Exception $Result.Exception
}
else 
{
    $Result.StatusCode  
} 

#endregion