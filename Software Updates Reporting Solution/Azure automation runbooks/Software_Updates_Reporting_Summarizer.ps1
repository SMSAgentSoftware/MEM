#####################################################################################
## Azure automation runbook that gathers additional WU related info from Microsoft ##
## and summarizes the current data for compliance reporting                        ##
#####################################################################################


##################
## AUTHENTICATE ##
##################
#region Authentication
$ResourceGroupName = "<ResourceGroupName>" # Name of the resource group containing your log analytics workspace
$WorkspaceName = "<WorkspaceName>" # The log analytics workspace name
$WorkspaceID = "<WorkspaceID>" # The WorkspaceID of the Log Analytics workspace
$PrimaryKey = "<PrimaryKey>" # The primary key of the log analytics workspace

$script:Destination = "$env:TEMP"
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
        #"https://aka.ms/WindowsUpdateHistory" # 2023-10-20 - MS broke this URL and it redirects to W11 not W10
        "https://support.microsoft.com/en-gb/topic/windows-10-update-history-7dd3071a-3906-fa2c-c342-f7f86728a6e3"
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

# Function to output a datatable containing the latest updates for each W10 version
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

# Function to get Windows update error codes from MS
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
    $headers = $htmlarray | Select-String -SimpleMatch "<h2 " | Where {$_ -match "error" -or $_ -match "success"} 
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
    $headers = $htmlarray | Select-String -SimpleMatch "<h2 id=""0x"
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
    
    #######################
    ## REMOVE DUPLICATES ##
    #######################
    # No need for duplicated error codes.
    [array]$Duplicates = $ErrorCodeTable | 
        Group-Object -Property ErrorCode,Category -NoElement | 
        Where {$_.Count -ge 2} | 
        Select -ExpandProperty Name
    If ($Duplicates.Count -ge 1)
    {
        foreach ($Duplicate in $Duplicates)
        {
            $ErrorCode = $Duplicate.Split(',')[0]
            $Rows = $ErrorCodeTable.Select("ErrorCode='$ErrorCode'")
            
            foreach ($Row in $Rows)
            {
                If ($ErrorCodeTable.Select("ErrorCode='$ErrorCode'").Count -gt 1)
                {
                    do 
                    {
                        $ErrorCodeTable.Rows.Remove($Row)
                    }
                    until ($ErrorCodeTable.Select("ErrorCode='$ErrorCode'").Count -eq 1)
                }
            }
            
        }
    }
    
    Return $ErrorCodeTable
}

# Function to get Windows setup error codes from MS
Function Get-WindowsSetupErrorCodes {
    $ProgressPreference = 'SilentlyContinue'
    $URL = "https://learn.microsoft.com/en-us/troubleshoot/windows-client/deployment/windows-10-upgrade-resolution-procedures"
    $tempFile = [System.IO.Path]::GetTempFileName() 
    Invoke-WebRequest -URI $URL -OutFile $tempFile -UseBasicParsing 
    $htmlarray = Get-Content $tempFile -ReadCount 0
    [System.IO.File]::Delete($tempFile)

    $headers = $htmlarray | Select-String -SimpleMatch "<h2 " | Where {$_ -notmatch "More information" -and $_ -notmatch "In this" -and $_ -notmatch "Additional resources" -and $_ -notmatch "Data collection" -and $_ -notmatch "Feedback"}
    $dataCells = $htmlarray | Select-String -SimpleMatch "<td>", "<td "

    $ErrorCodeTable = [System.Data.DataTable]::new()
    $ErrorCodeTable.Columns.AddRange(@("ErrorCode","ExtendedCode","Description","Message","Category"))

    # ref: https://www.dotnetperls.com/remove-html-tags
    function Remove-HTMLFromString {
        Param ($htmlstring)
        $array = [Char[]]::new($htmlstring.length)
        [int]$arrayIndex = 0
        [bool]$Inside = $false

        for ($i=0; $i -lt $htmlstring.Length;$i++)
        {
            [char]$let = $htmlstring[$i]
            if ($let -eq "<")
            {
                $Inside = $true
                Continue
            }
            if ($let -eq ">")
            {
                $Inside = $false
                Continue
            }
            if (!($Inside))
            {
                $array[$arrayIndex] = $let
                $arrayIndex++
            }
        }
        return [System.String]::new([char[]]$array,[int]0,[int]$arrayIndex)
    }

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
                $WebErrorCode = "$($cells[$t].ToString().Replace('<br>',',').Split('>').Split('<')[2])"
                $WebErrorCodeArray = $WebErrorCode.Split(',')
                If ($WebErrorCodeArray.Count -gt 1)
                {
                    1..($WebErrorCodeArray.Count) | foreach {
                        $Row = $ErrorCodeTable.NewRow()
                        $Cell = $WebErrorCodeArray[($_ -1)]
                        "ErrorCode","ExtendedCode","Message","Description" | foreach {
                            If ($_ -eq "ErrorCode")
                            {
                                If ($Cell -match "-")
                                {
                                    $CellSplit = $Cell.Split('-')[0].Trim()
                                    If ($CellSplit.StartsWith('80'))
                                    {
                                        $CellSplit = "0x$($CellSplit)"
                                    }
                                    $Row["$_"] = $CellSplit
                                }
                                else 
                                {
                                    $Row["$_"] = $Cell
                                }
                            }
                            ElseIf ($_ -eq "ExtendedCode")
                            {
                                If ($Cell -match "-")
                                {
                                    $CellSplit = $Cell.Split('-')[1].Trim()
                                    $Row["$_"] = $CellSplit
                                }
                                else 
                                {
                                    $Row["$_"] = $null
                                }
                            }
                            else 
                            {
                                $Row["$_"] = Remove-HTMLFromString "$($cells[$t].ToString().Replace('<br>',[Environment]::NewLine).Replace('Â','').Replace('&quot;','""').Replace('&gt;','-'))"
                            }
                            If ($_ -ne "ExtendedCode"){$t ++}
                        }
                        $Row["Category"] = "$($header.ToString().Split('>').Split('<')[2])"
                        [void]$ErrorCodeTable.Rows.Add($Row)
                        1..3 | foreach {$t --}
                    }
                    1..($WebErrorCodeArray.Count) | foreach {$t++}
                }
                else 
                {
                    $Row = $ErrorCodeTable.NewRow() 
                    "ErrorCode","ExtendedCode","Message","Description" | foreach {
                        $Cell = "$($cells[$t].ToString().Replace('<br>',[Environment]::NewLine).Split('>').Split('<')[2])"
                        If ($_ -eq "ErrorCode")
                        {
                            If ($cell -match "-")
                            {
                                $CellSplit = "$($cell.Split('-')[0].Trim())"
                                If ($CellSplit.StartsWith('80'))
                                {
                                    $CellSplit = "0x$($CellSplit)"
                                }
                                $Row["$_"] = $CellSplit
                            }
                            else 
                            {
                                $Row["$_"] = $cell
                            }
                        }
                        ElseIf ($_ -eq "ExtendedCode")
                        {
                            If ($cell -match "-")
                            {
                                $CellSplit = "$($cell.Split('-')[1].Trim())"
                                $Row["$_"] = $CellSplit
                            }
                            else 
                            {
                                $Row["$_"] = $null
                            }
                        }
                        else 
                        {
                            $Row["$_"] = Remove-HTMLFromString "$($cells[$t].ToString().Replace('<br>',[Environment]::NewLine).Replace('Â','').Replace('&quot;','""').Replace('&gt;','-'))"
                        }
                        
                        If ($_ -ne "ErrorCode"){$t ++}
                    }
                    $Row["Category"] = "$($header.ToString().Split('>').Split('<')[2])"
                    [void]$ErrorCodeTable.Rows.Add($Row)
                }                     
            }
            until ($t -ge ($totalCells -1))
            $i ++
        }
    }
    until ($i -ge $headers.count)

    # Remove duplicates
    $Rows = $ErrorCodeTable.Select("ErrorCode = '0XC1900200'") 
    $ErrorCodeTable.Rows.Remove($Rows[2])
    $ErrorCodeTable.Rows.Remove($Rows[-1])

    $Rows = $ErrorCodeTable.Select("ErrorCode = '0xC190020e'") 
    $ErrorCodeTable.Rows.Remove($Rows[-1])

    $Rows = $ErrorCodeTable.Select("ErrorCode = '0xC1900209'") 
    $ErrorCodeTable.Rows.Remove($Rows[-1])

    $Rows = $ErrorCodeTable.Select("ErrorCode = '0xC1900201'") 
    $ErrorCodeTable.Rows.Remove($Rows[-1])

    $Rows = $ErrorCodeTable.Select("ErrorCode = '0xC1900107'") 
    $ErrorCodeTable.Rows.Remove($Rows[-1])

    $Rows = $ErrorCodeTable.Select("ErrorCode = '0xC1900101' and ExtendedCode = '0x2000c'") 
    $ErrorCodeTable.Rows.Remove($Rows[-1])

    return $ErrorCodeTable
}
#endregion

#################################
## UPDATE WU ERROR CODES TABLE ##
#################################
#region UpdateWUErrorCodes
# This runs twice a month just to keep the data from ageing past the data retention period in the LA workspace
# Remove the surrounding IF statement for a first-time run, so you have some data right away
#If ([DateTime]::UtcNow.Day -eq 7 -or [DateTime]::UtcNow.Day -eq 21)
#{
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

    $WindowsSetupErrorCodes = Get-WindowsSetupErrorCodes
    $Table = [System.Data.DataTable]::new()
    ($WindowsSetupErrorCodes[0] | Get-Member -MemberType Property).Name | foreach {
        [void]$Table.Columns.Add($_)
    }
    foreach ($row in $WindowsSetupErrorCodes)
    {
        $Table.ImportRow($row)
    }
    # Post the JSON to LA workspace
    $Json = $Table.Rows | Select ErrorCode,ExtendedCode,Description,Message,Category | ConvertTo-Json -Compress
    $Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_WindowsSetupErrorCodes"
    If ($Result.GetType().Name -eq "ErrorRecord")
    {
        Write-Error -Exception $Result.Exception
    }
    else 
    {
        $Result.StatusCode  
    } 
#}
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

# Microsoft boo-boo: 18362.1916 should be 18363.1916
$Correction1 = [Linq.Enumerable]::FirstOrDefault($UpdateHistoryTable.Select("OSBuild='18362.1916'"))
If ($null -ne $Correction1)
{
    $Correction1.OSBuild = "18363.1916"
    $Correction1.OSBaseBuild = "18363"
    $Correction1.OSVersion = "1909"
}

# Microsoft boo-boo: 18362.1977 should be 18363.1977
$Correction2 = [Linq.Enumerable]::FirstOrDefault($UpdateHistoryTable.Select("OSBuild='18362.1977'"))
If ($null -ne $Correction2)
{
    $Correction2.OSBuild = "18363.1977"
    $Correction2.OSBaseBuild = "18363"
    $Correction2.OSVersion = "1909"
}

# Latest update table must be created after the above changes
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

###########################
## EXECUTE DEVICES QUERY ##
###########################
#region DeviceQuery
$Query = @"
let DevicesBase = SU_DeviceInfo_CL 
| where isnotnull(InventoryDate_t)
| where DisplayVersion_s != `"Dev`"
| summarize arg_max(InventoryDate_t,*) by IntuneDeviceID_g
| join kind = leftouter (SU_OSUpdateHistory_CL 
    | summarize arg_max(TimeGenerated,*) by OSBuild_s
)on `$left.CurrentPatchLevel_s == `$right.OSBuild_s;
let DevicesWithLatestUpdates = DevicesBase
| join kind=leftouter (SU_OSLatestUpdates_CL 
    | top-nested 1 of TimeGenerated by temp=max(TimeGenerated),
        top-nested of Windows_Release=Windows_Release_s by temp1=max(1),
        top-nested of OSBaseBuild=OSBaseBuild_d by temp2=max(1),
        top-nested of OSVersion=OSVersion_s by temp3=max(1),
        top-nested of LatestUpdate=LatestUpdate_s by temp4=max(1),
        top-nested of LatestUpdateKB=LatestUpdate_KB_s by temp5=max(1),
        top-nested of LatestUpdateReleaseDate=LatestUpdate_ReleaseDate_s by temp5a=max(1),
        top-nested of LatestRegularUpdate=LatestRegularUpdate_s by temp6=max(1),
        top-nested of LatestRegularUpdateKB=LatestRegularUpdate_KB_s by temp7=max(1),
        top-nested of LatestRegularUpdateReleaseDate=LatestRegularUpdate_ReleaseDate_s by temp7a=max(1),
        top-nested of LatestPreviewUpdate=LatestPreviewUpdate_s by temp8=max(1),
        top-nested of LatestPreviewUpdateKB=LatestPreviewUpdate_KB_s by temp9=max(1),
        top-nested of LatestPreviewUpdateReleaseDate=LatestPreviewUpdate_ReleaseDate_s by temp9a=max(1),
        top-nested of LatestOutofBandUpdate=LatestOutofBandUpdate_s by temp10=max(1),
        top-nested of LatestOutofBandUpdateKB=LatestOutofBandUpdate_KB_s by temp11=max(1),
        top-nested of LatestOutofBandUpdateReleaseDate=LatestOutofBandUpdate_ReleaseDate_s by temp11a=max(1),
        top-nested of LatestRegularUpdateLess1=LatestRegularUpdateLess1_s by temp12=max(1),
        top-nested of LatestRegularUpdateLess1KB=LatestRegularUpdateLess1_KB_s by temp13=max(1),
        top-nested of LatestRegularUpdateLess1ReleaseDate=LatestRegularUpdateLess1_ReleaseDate_s by temp13a=max(1),
        top-nested of LatestPreviewUpdateLess1=LatestPreviewUpdateLess1_s by temp14=max(1),
        top-nested of LatestPreviewUpdateLess1KB=LatestPreviewUpdateLess1_KB_s by temp15=max(1),
        top-nested of LatestPreviewUpdateLess1ReleaseDate=LatestPreviewUpdateLess1_ReleaseDate_s by temp15a=max(1),
        top-nested of LatestOutofBandUpdateLess1=LatestOutofBandUpdateLess1_s by temp16=max(1),
        top-nested of LatestOutofBandUpdateLess1KB=LatestOutofBandUpdateLess1_KB_s by temp17=max(1),
        top-nested of LatestOutofBandUpdateLess1ReleaseDate=LatestOutofBandUpdateLess1_ReleaseDate_s by temp17a=max(1),
        top-nested of LatestRegularUpdateLess2=LatestRegularUpdateLess2_s by temp18=max(1),
        top-nested of LatestRegularUpdateLess2KB=LatestRegularUpdateLess2_KB_s by temp19=max(1),
        top-nested of LatestRegularUpdateLess2ReleaseDate=LatestRegularUpdateLess2_ReleaseDate_s by temp19a=max(1),
        top-nested of LatestPreviewUpdateLess2=LatestPreviewUpdateLess2_s by temp20=max(1),
        top-nested of LatestPreviewUpdateLess2KB=LatestPreviewUpdateLess2_KB_s by temp21=max(1),
        top-nested of LatestPreviewUpdateLess2ReleaseDate=LatestPreviewUpdateLess2_ReleaseDate_s by temp21a=max(1),
        top-nested of LatestOutofBandUpdateLess2=LatestOutofBandUpdateLess2_s by temp22=max(1),
        top-nested of LatestOutofBandUpdateLess2KB=LatestOutofBandUpdateLess2_KB_s by temp23=max(1),
        top-nested of LatestOutofBandUpdateLess2ReleaseDate=LatestOutofBandUpdateLess2_ReleaseDate_s by temp23a=max(1),
        top-nested of LatestUpdateType=LatestUpdateType_s by temp24=max(1)
    | project-away temp*
    | order by Windows_Release,OSBaseBuild desc 
)on `$left.Windows_Release_s == `$right.Windows_Release and `$left.OSVersion_s == `$right.OSVersion;
let DevicesWithUpdateLog = DevicesWithLatestUpdates
| join kind=leftouter (SU_UpdateLog_CL
    | where UpdateType_s == `"Windows cumulative update`"
    | top-nested of IntuneDeviceID_g by temp99=max(1),
        top-nested 1 of InventoryDate_t by temp=max(InventoryDate_t),
        top-nested of KB=KB_s by temp1=max(1),
        top-nested of EventId=EventId_d by temp2=max(1),
        top-nested of KeyWord1=KeyWord1_s by temp3=max(1),
        top-nested of KeyWord2=KeyWord2_s by temp4=max(1),
        top-nested of RebootRequired=column_ifexists(`"RebootRequired_s`",`"`") by temp5=max(1),
        top-nested of ServiceGuid=ServiceGuid_g by temp6=max(1),
        top-nested of ServiceName=ServiceName_s by temp7=max(1),
        top-nested of TimeCreated=TimeCreated_t by temp8=max(1),
        top-nested of UpdateName=UpdateName_s by temp9=max(1),
        top-nested of UpdateType=UpdateType_s by temp10=max(1),
        top-nested of WindowsVersion=WindowsVersion_s by temp11=max(1),
        top-nested of WindowsDisplayVersion=WindowsDisplayVersion_s by temp12=max(1)
    | project-away temp*
)on IntuneDeviceID_g and `$left.LatestRegularUpdateKB == `$right.KB and `$left.Windows_Release_s == `$right.WindowsVersion and `$left.DisplayVersion_s == `$right.WindowsDisplayVersion;
let DevicesWithWUClientInfo = DevicesWithUpdateLog
| join kind=leftouter (SU_WUClientInfo_CL
    | summarize arg_max(InventoryDate_t,*) by IntuneDeviceID_g
)on IntuneDeviceID_g;
let Devices = DevicesWithWUClientInfo
    | join kind=leftouter (SU_WUPolicyState_CL
    | summarize arg_max(InventoryDate_t,*) by IntuneDeviceID_g
    | project  QualityUpdatesDeferralInDays_d, FeatureUpdatesDeferralInDays_d, FeatureUpdatesPaused_d=column_ifexists(`"FeatureUpdatesPaused_d`",real(null)), QualityUpdatesPaused_d=column_ifexists(`"QualityUpdatesPaused_d`",real(null)), FeatureUpdatePausePeriodInDays_d=column_ifexists(`"FeatureUpdatePausePeriodInDays_d`",real(null)), QualityUpdatePausePeriodInDays_d=column_ifexists(`"QualityUpdatePausePeriodInDays_d`",real(null)), PauseFeatureUpdatesStartTime_t=column_ifexists(`"PauseFeatureUpdatesStartTime_t`",datetime(null)), PauseQualityUpdatesStartTime_t=column_ifexists(`"PauseQualityUpdatesStartTime_t`",datetime(null)), PauseFeatureUpdatesEndTime_t=column_ifexists(`"PauseFeatureUpdatesEndTime_t`",datetime(null)), PauseQualityUpdatesEndTime_t=column_ifexists(`"PauseQualityUpdatesEndTime_t`",datetime(null)), IntuneDeviceID_g
)on IntuneDeviceID_g;
Devices
| extend IsLatestOSBuild = case(
    OSBuild_s == LatestUpdate,`"Yes`",`"No`"
)
| extend IsLatestRegularOSBuild = case(
    OSBuild_s == LatestRegularUpdate,`"Yes`",`"No`"
)
| extend IsLatestPreviewOSBuild = case(
    OSBuild_s == LatestPreviewUpdate,`"Yes`",`"No`"
)
| extend IsLatestOutofBandOSBuild = case(
    OSBuild_s == LatestOutofBandUpdate,`"Yes`",`"No`"
)
| extend CurrentPatchLevelAgeinDays = datetime_diff('day',now(),ReleaseDate_t)
| extend LatestRegularUpdateName = iff(isnotempty(LatestRegularUpdateKB),strcat(Windows_Release_s,`" `",OSVersion_s,`"  -  `",LatestRegularUpdateKB, `" [`",LatestRegularUpdate,`"] [Security 'B'] `",iff(isempty(LatestRegularUpdateReleaseDate),`"`",format_datetime(todatetime(LatestRegularUpdateReleaseDate),'yyyy-MM-dd'))),`"`")
| extend LatestPreviewUpdateName = iff(isnotempty(LatestPreviewUpdateKB),strcat(Windows_Release_s,`" `",OSVersion_s,`"  -  `",LatestPreviewUpdateKB, `" [`",LatestPreviewUpdate,`"] [Non-Security Preview] `",iff(isempty(LatestPreviewUpdateReleaseDate),`"`",format_datetime(todatetime(LatestPreviewUpdateReleaseDate),'yyyy-MM-dd'))),`"`")
| extend LatestOutofBandUpdateName = iff(isnotempty(LatestOutofBandUpdateKB),strcat(Windows_Release_s,`" `",OSVersion_s,`"  -  `",LatestOutofBandUpdateKB, `" [`",LatestOutofBandUpdate,`"] [Out-of-band] `",iff(isempty(LatestOutofBandUpdateReleaseDate),`"`",format_datetime(todatetime(LatestOutofBandUpdateReleaseDate),'yyyy-MM-dd'))),`"`")
| extend LatestRegularUpdateLess1Name = iff(isnotempty(LatestRegularUpdateLess1KB),strcat(Windows_Release_s,`" `",OSVersion_s,`"  -  `",LatestRegularUpdateLess1KB, `" [`",LatestRegularUpdateLess1,`"] [Security 'B'] `",iff(isempty(LatestRegularUpdateLess1ReleaseDate),`"`",format_datetime(todatetime(LatestRegularUpdateLess1ReleaseDate),'yyyy-MM-dd'))),`"`")
| extend LatestPreviewUpdateLess1Name = iff(isnotempty(LatestPreviewUpdateLess1KB),strcat(Windows_Release_s,`" `",OSVersion_s,`"  -  `",LatestPreviewUpdateLess1KB, `" [`",LatestPreviewUpdateLess1,`"] [Non-Security Preview] `",iff(isempty(LatestPreviewUpdateLess1ReleaseDate),`"`",format_datetime(todatetime(LatestPreviewUpdateLess1ReleaseDate),'yyyy-MM-dd'))),`"`")
| extend LatestOutofBandUpdateLess1Name = iff(isnotempty(LatestOutofBandUpdateLess1KB),strcat(Windows_Release_s,`" `",OSVersion_s,`"  -  `",LatestOutofBandUpdateLess1KB, `" [`",LatestOutofBandUpdateLess1,`"]  [Out-of-band] `",iff(isempty(LatestOutofBandUpdateLess1ReleaseDate),`"`",format_datetime(todatetime(LatestOutofBandUpdateLess1ReleaseDate),'yyyy-MM-dd'))),`"`")
| extend LatestRegularUpdateLess2Name = iff(isnotempty(LatestRegularUpdateLess2KB),strcat(Windows_Release_s,`" `",OSVersion_s,`"  -  `",LatestRegularUpdateLess2KB, `" [`",LatestRegularUpdateLess2,`"] [Security 'B'] `",iff(isempty(LatestRegularUpdateLess2ReleaseDate),`"`",format_datetime(todatetime(LatestRegularUpdateLess2ReleaseDate),'yyyy-MM-dd'))),`"`")
| extend LatestPreviewUpdateLess2Name = iff(isnotempty(LatestPreviewUpdateLess2KB),strcat(Windows_Release_s,`" `",OSVersion_s,`"  -  `",LatestPreviewUpdateLess2KB, `" [`",LatestPreviewUpdateLess2,`"] [Non-Security Preview] `",iff(isempty(LatestPreviewUpdateLess2ReleaseDate),`"`",format_datetime(todatetime(LatestPreviewUpdateLess2ReleaseDate),'yyyy-MM-dd'))),`"`")
| extend LatestOutofBandUpdateLess2Name = iff(isnotempty(LatestOutofBandUpdateLess2KB),strcat(Windows_Release_s,`" `",OSVersion_s,`"  -  `",LatestOutofBandUpdateLess2KB, `" [`",LatestOutofBandUpdateLess2,`"]  [Out-of-band] `",iff(isempty(LatestOutofBandUpdateLess2ReleaseDate),`"`",format_datetime(todatetime(LatestOutofBandUpdateLess2ReleaseDate),'yyyy-MM-dd'))),`"`")
| project 
    InventoryDate=InventoryDate_t,
    ComputerName=ComputerName_s,
    LatestInventoryType=LatestInventoryType_s,
    LatestDeltaInventory=LatestDeltaInventory_t,
    LatestFullInventory=LatestFullInventory_t,
    InventoryExecutionDuration=InventoryExecutionDuration_d,
    AADDeviceID=AADDeviceID_g,
    IntuneDeviceID=IntuneDeviceID_g,
    LastSyncTime=LastSyncTime_t,
    CurrentUser=CurrentUser_s,
    FriendlyOSName=FriendlyOSName_s,
    FullBuildNmber=FullBuildNmber_s,
    CurrentBuildNumber=CurrentBuildNumber_s,
    EditionID=EditionID_s,
    Manufacturer=Manufacturer_s,
    Model=Model_s,
    DisplayVersion=DisplayVersion_s,
    ProductName=ProductName_s,
    CurrentPatchLevel=CurrentPatchLevel_s,
    OSBuild=OSBuild_s,
    Windows_Release=Windows_Release_s,
    ReleaseDate=ReleaseDate_t,
    KB=KB_s,
    OSBaseBuild=OSBaseBuild_d,
    OSRevisionNumber=OSRevisionNumber_d,
    OSVersion=OSVersion_s,
    PatchType=Type_s,
    PatchStatusDate=TimeCreated,
    UpdateActivity=KeyWord1,
    UpdateStatus=KeyWord2,
    EventId=EventId,
    UpdateName=UpdateName,
    UpdateType=UpdateType,
    ServiceGuid=ServiceGuid,
    ServiceName=ServiceName,
    PatchRebootRequired=RebootRequired,
    EngageReminderLastShownTime=column_ifexists(`"EngageReminderLastShownTime_t`",datetime(null)),
    ScheduledRebootTime=column_ifexists(`"ScheduledRebootTime_t`",datetime(null)),
    PendingRebootStartTime=column_ifexists(`"PendingRebootStartTime_t`",datetime(null)),
    AutoUpdateStatus=AutoUpdateStatus_s,
    WURebootRequired=column_ifexists(`"RebootRequired_s`",`"`"),
    NoAutoRebootWithLoggedOnUsers=NoAutoRebootWithLoggedOnUsers_s,
    WUServiceStartupType=WUServiceStartupType_s,
    LatestRegularUpdate,
    LatestRegularUpdateKB,
    LatestRegularUpdateReleaseDate,
    LatestRegularUpdateName,
    LatestPreviewUpdate,
    LatestPreviewUpdateKB,
    LatestPreviewUpdateReleaseDate,
    LatestPreviewUpdateName,
    LatestOutofBandUpdate,
    LatestOutofBandUpdateKB,
    LatestOutofBandUpdateReleaseDate,
    LatestOutofBandUpdateName,
    LatestUpdateType,
    LatestRegularUpdateLess1,
    LatestRegularUpdateLess1KB,
    LatestRegularUpdateLess1ReleaseDate,
    LatestRegularUpdateLess1Name,
    LatestPreviewUpdateLess1,
    LatestPreviewUpdateLess1KB,
    LatestPreviewUpdateLess1ReleaseDate,
    LatestPreviewUpdateLess1Name,
    LatestOutofBandUpdateLess1,
    LatestOutofBandUpdateLess1KB,
    LatestOutofBandUpdateLess1ReleaseDate,
    LatestOutofBandUpdateLess1Name,
    LatestRegularUpdateLess2,
    LatestRegularUpdateLess2KB,
    LatestRegularUpdateLess2ReleaseDate,
    LatestRegularUpdateLess2Name,
    LatestPreviewUpdateLess2,
    LatestPreviewUpdateLess2KB,
    LatestPreviewUpdateLess2ReleaseDate,
    LatestPreviewUpdateLess2Name,
    LatestOutofBandUpdateLess2,
    LatestOutofBandUpdateLess2KB,
    LatestOutofBandUpdateLess2ReleaseDate,
    LatestOutofBandUpdateLess2Name,
    IsLatestOSBuild,
    IsLatestRegularOSBuild,
    IsLatestOutofBandOSBuild,
    IsLatestPreviewOSBuild,
    QualityUpdatesDeferralInDays=QualityUpdatesDeferralInDays_d,
    FeatureUpdatesDeferralInDays=FeatureUpdatesDeferralInDays_d,
    FeatureUpdatesPaused=FeatureUpdatesPaused_d,
    QualityUpdatesPaused=QualityUpdatesPaused_d,
    FeatureUpdatePausePeriodInDays=FeatureUpdatePausePeriodInDays_d,
    QualityUpdatePausePeriodInDays=QualityUpdatePausePeriodInDays_d,
    PauseFeatureUpdatesStartTime=PauseFeatureUpdatesStartTime_t,
    PauseQualityUpdatesStartTime=PauseQualityUpdatesStartTime_t,
    PauseFeatureUpdatesEndTime=PauseFeatureUpdatesEndTime_t,
    PauseQualityUpdatesEndTime=PauseQualityUpdatesEndTime_t,
    CurrentPatchLevelAgeinDays
| order by ComputerName asc
"@
try 
{
    $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -Timespan (New-TimeSpan -Days 30) -IncludeStatistics -ErrorAction Stop
}
catch 
{
    Write-Error "Invocation of the Log Analytics query failed: $($_.Exception.Message)"
    Write-Output "Let's try the LA query again..."
    try 
    {
        $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -Timespan (New-TimeSpan -Days 30) -IncludeStatistics -ErrorAction Stop
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

# This is to mitigate an issue where sometimes the query results returned are not as expected, with large numbers of update data not populated.
$LatestRegularUpdateCount = ($Result.results.Where({$_.LatestRegularUpdate.Length -eq 0})).Count 
If ($LatestRegularUpdateCount -ge 500)
{
    Write-Warning "$LatestRegularUpdateCount results were returned with no data for Latest regular update. Probably this query didn't execute properly. Re-running."
    Start-Sleep -Seconds 20
    try 
    {
        $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -Timespan (New-TimeSpan -Days 30) -IncludeStatistics -ErrorAction Stop
    }
    catch 
    {
        Write-Error "Invocation of the Log Analytics query failed: $($_.Exception.Message)"
        Write-Output "Let's try the LA query again..."
        try 
        {
            $Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -Timespan (New-TimeSpan -Days 30) -IncludeStatistics -ErrorAction Stop
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
        Write-Host "LA Query Stats"
        Write-Host "=============="
        Write-Host "CPU time (hh:mm:ss): $CPUTime"
        Write-Host "Row count: $($TableStats.tableRowCount)"
        Write-Host "Table size (MB): $($TableStats.tableSize / 1MB)" 
    }

    $LatestRegularUpdateCount = ($Result.results.Where({$_.LatestRegularUpdate.Length -eq 0})).Count
    If ($LatestRegularUpdateCount -ge 500)
    {
        throw "The Log Analytics query ran twice and failed to return the expected results. Giving up for this run."
    }
}
#endregion

###############################
## CONVERT DATA TO DATATABLE ##
###############################
#region ConvertToTable
$iDevices = $Result.Results
$DevicesArray = [System.Linq.Enumerable]::ToArray($iDevices)
$DevicesDatatable = [System.Data.DataTable]::new()
$ColumnNames = ($iDevices | Get-Member -MemberType NoteProperty).Name
foreach ($ColumnName in $ColumnNames) {
    [void]$DevicesDatatable.Columns.AddRange($ColumnName)
}
foreach ($item in $DevicesArray)
{
    $NewRow = $DevicesDatatable.NewRow()
    foreach ($ColumnName in $ColumnNames)
    {
        $NewRow["$ColumnName"] = $Item.$ColumnName
    }
    [void]$DevicesDatatable.Rows.Add($NewRow)
}
# Add some additional columns for calculated values
@(
    'ComplianceStateLatest'
    'ComplianceStateRegular'
    'ComplianceStatePreview'
    'ComplianceStateOutofBand'
    'ComplianceStateRegularLess1'
    'ComplianceStatePreviewLess1'
    'ComplianceStateOutofBandLess1'
    'ComplianceStateRegularLess2'
    'ComplianceStatePreviewLess2'
    'ComplianceStateOutofBandLess2'
    'OSVersionSupportStatus'
    'SupportDaysRemaining'
    'LatestRegularUpdateStatus'
    'SummarizationTime'
) | foreach {
    [void]$DevicesDatatable.Columns.Add($_)
}
#endregion

#############################################
## ADJUSTMENTS FOR RTM OR PREVIEW RELEASES ##
#############################################
#region RTM
foreach ($Row in $DevicesDatatable)
{
    # For RTM or preview releases, the OSBuild value will be empty as there is no match against any update in the UpdateHistory table.
    # Therefore we need to manually populate some values here.
    If ($Row["OSBuild"].Length -eq 0)
    {
        $Row["PatchType"] = "RTM or Preview"
        $Row["IsLatestOSBuild"] = "N/A"
        $Row["IsLatestRegularOSBuild"] = "N/A"
        $Row["IsLatestOutofBandOSBuild"] = "N/A"
        $Row["IsLatestPreviewOSBuild"] = "N/A"
        $ProductNameSplit = $Row["ProductName"].Split()
        $Row["Windows_Release"] = "$($ProductNameSplit[0]) $($ProductNameSplit[1])"
        $Row["OSBaseBuild"] = $Row["CurrentBuildNumber"]
        $Row["OSRevisionNumber"] = $Row["CurrentPatchLevel"].Split('.')[1]
        $Row["OSVersion"] = $Row["DisplayVersion"]

        $QueryString = "[Windows Release]='$($Row['Windows_Release'])' and OSBaseBuild='$($Row['OSBaseBuild'])'"
        $LatestUpdates = [Linq.Enumerable]::FirstOrDefault($LatestUpdateTable.Select($QueryString))
        If ($null -ne $LatestUpdates)
        {
            $Row["LatestRegularUpdate"] = $LatestUpdates.LatestRegularUpdate
            $Row["LatestRegularUpdateKB"] = $LatestUpdates.LatestRegularUpdate_KB
            $Row["LatestRegularUpdateReleaseDate"] = $LatestUpdates.LatestRegularUpdate_ReleaseDate
            If ($Row["LatestRegularUpdateKB"].Length -ge 1 -and $Row["LatestRegularUpdateReleaseDate"].GetType() -ne [System.DBNull])
            {
                $Row["LatestRegularUpdateName"] = "$($Row["Windows_Release"]) $($Row["OSVersion"])  -  $($Row["LatestRegularUpdateKB"]) [$($Row["LatestRegularUpdate"])] [Security 'B'] $(Get-Date $Row["LatestRegularUpdateReleaseDate"] -Format 'yyyy-MM-dd')"
            }
            $Row["LatestPreviewUpdate"] = $LatestUpdates.LatestPreviewUpdate
            $Row["LatestPreviewUpdateKB"] = $LatestUpdates.LatestPreviewUpdate_KB
            $Row["LatestPreviewUpdateReleaseDate"] = $LatestUpdates.LatestPreviewUpdate_ReleaseDate
            If ($Row["LatestPreviewUpdateKB"].Length -ge 1 -and $Row["LatestPreviewUpdateReleaseDate"].GetType() -ne [System.DBNull])
            {
                $Row["LatestPreviewUpdateName"] = "$($Row["Windows_Release"]) $($Row["OSVersion"])  -  $($Row["LatestPreviewUpdateKB"]) [$($Row["LatestPreviewUpdate"])] [Non-Security Preview] $(Get-Date $Row["LatestPreviewUpdateReleaseDate"] -Format 'yyyy-MM-dd')"
            }
            $Row["LatestOutofBandUpdate"] = $LatestUpdates.LatestOutofBandUpdate
            $Row["LatestOutofBandUpdateKB"] = $LatestUpdates.LatestOutofBandUpdate_KB
            $Row["LatestOutofBandUpdateReleaseDate"] = $LatestUpdates.LatestOutofBandUpdate_ReleaseDate
            If ($Row["LatestOutofBandUpdateKB"].Length -ge 1 -and $Row["LatestOutofBandUpdateReleaseDate"].GetType() -ne [System.DBNull])
            {
                $Row["LatestOutofBandUpdateName"] = "$($Row["Windows_Release"]) $($Row["OSVersion"])  -  $($Row["LatestOutofBandUpdateKB"]) [$($Row["LatestOutofBandUpdate"])] [Out of Band] $(Get-Date $Row["LatestOutofBandUpdateReleaseDate"] -Format 'yyyy-MM-dd')"
            }
            $Row["LatestRegularUpdateLess1"] = $LatestUpdates.LatestRegularUpdateLess1
            $Row["LatestRegularUpdateLess1KB"] = $LatestUpdates.LatestRegularUpdateLess1_KB
            $Row["LatestRegularUpdateLess1ReleaseDate"] = $LatestUpdates.LatestRegularUpdateLess1_ReleaseDate
            If ($Row["LatestRegularUpdateLess1KB"].Length -ge 1 -and $Row["LatestRegularUpdateLess1ReleaseDate"].GetType() -ne [System.DBNull])
            {
                $Row["LatestRegularUpdateLess1Name"] = "$($Row["Windows_Release"]) $($Row["OSVersion"])  -  $($Row["LatestRegularUpdateLess1KB"]) [$($Row["LatestRegularUpdateLess1"])] [Security 'B'] $(Get-Date $Row["LatestRegularUpdateLess1ReleaseDate"] -Format 'yyyy-MM-dd')"
            }
            $Row["LatestPreviewUpdateLess1"] = $LatestUpdates.LatestPreviewUpdateLess1
            $Row["LatestPreviewUpdateLess1KB"] = $LatestUpdates.LatestPreviewUpdateLess1_KB
            $Row["LatestPreviewUpdateLess1ReleaseDate"] = $LatestUpdates.LatestPreviewUpdateLess1_ReleaseDate
            If ($Row["LatestPreviewUpdateLess1KB"].Length -ge 1 -and $Row["LatestPreviewUpdateLess1ReleaseDate"].GetType() -ne [System.DBNull])
            {
                $Row["LatestPreviewUpdateLess1Name"] = "$($Row["Windows_Release"]) $($Row["OSVersion"])  -  $($Row["LatestPreviewUpdateLess1KB"]) [$($Row["LatestPreviewUpdateLess1"])] [Non-Security Preview] $(Get-Date $Row["LatestPreviewUpdateLess1ReleaseDate"] -Format 'yyyy-MM-dd')"
            }
            $Row["LatestOutofBandUpdateLess1"] = $LatestUpdates.LatestOutofBandUpdateLess1
            $Row["LatestOutofBandUpdateLess1KB"] = $LatestUpdates.LatestOutofBandUpdateLess1_KB
            $Row["LatestOutofBandUpdateLess1ReleaseDate"] = $LatestUpdates.LatestOutofBandUpdateLess1_ReleaseDate
            If ($Row["LatestOutofBandUpdateLess1KB"].Length -ge 1 -and $Row["LatestOutofBandUpdateLess1ReleaseDate"].GetType() -ne [System.DBNull])
            {
                $Row["LatestOutofBandUpdateLess1Name"] = "$($Row["Windows_Release"]) $($Row["OSVersion"])  -  $($Row["LatestOutofBandUpdateLess1KB"]) [$($Row["LatestOutofBandUpdateLess1"])] [Out of Band] $(Get-Date $Row["LatestOutofBandUpdateLess1ReleaseDate"] -Format 'yyyy-MM-dd')"
            }
            $Row["LatestRegularUpdateLess2"] = $LatestUpdates.LatestRegularUpdateLess2
            $Row["LatestRegularUpdateLess2KB"] = $LatestUpdates.LatestRegularUpdateLess2_KB
            $Row["LatestRegularUpdateLess2ReleaseDate"] = $LatestUpdates.LatestRegularUpdateLess2_ReleaseDate
            If ($Row["LatestRegularUpdateLess2KB"].Length -ge 1 -and $Row["LatestRegularUpdateLess2ReleaseDate"].GetType() -ne [System.DBNull])
            {
                $Row["LatestRegularUpdateLess2Name"] = "$($Row["Windows_Release"]) $($Row["OSVersion"])  -  $($Row["LatestRegularUpdateLess2KB"]) [$($Row["LatestRegularUpdateLess2"])] [Security 'B'] $(Get-Date $Row["LatestRegularUpdateLess2ReleaseDate"] -Format 'yyyy-MM-dd')"
            }
            $Row["LatestPreviewUpdateLess2"] = $LatestUpdates.LatestPreviewUpdateLess2
            $Row["LatestPreviewUpdateLess2KB"] = $LatestUpdates.LatestPreviewUpdateLess2_KB
            $Row["LatestPreviewUpdateLess2ReleaseDate"] = $LatestUpdates.LatestPreviewUpdateLess2_ReleaseDate
            If ($Row["LatestPreviewUpdateLess2KB"].Length -ge 1 -and $Row["LatestPreviewUpdateLess2ReleaseDate"].GetType() -ne [System.DBNull])
            {
                $Row["LatestPreviewUpdateLess2Name"] = "$($Row["Windows_Release"]) $($Row["OSVersion"])  -  $($Row["LatestPreviewUpdateLess2KB"]) [$($Row["LatestPreviewUpdateLess2"])] [Non-Security Preview] $(Get-Date $Row["LatestPreviewUpdateLess2ReleaseDate"] -Format 'yyyy-MM-dd')"
            }
            $Row["LatestOutofBandUpdateLess2"] = $LatestUpdates.LatestOutofBandUpdateLess2
            $Row["LatestOutofBandUpdateLess2KB"] = $LatestUpdates.LatestOutofBandUpdateLess2_KB
            $Row["LatestOutofBandUpdateLess2ReleaseDate"] = $LatestUpdates.LatestOutofBandUpdateLess2_ReleaseDate
            If ($Row["LatestOutofBandUpdateLess2KB"].Length -ge 1 -and $Row["LatestOutofBandUpdateLess2ReleaseDate"].GetType() -ne [System.DBNull])
            {
                $Row["LatestOutofBandUpdateLess2Name"] = "$($Row["Windows_Release"]) $($Row["OSVersion"])  -  $($Row["LatestOutofBandUpdateLess2KB"]) [$($Row["LatestOutofBandUpdateLess2"])] [Out of Band] $(Get-Date $Row["LatestOutofBandUpdateLess2ReleaseDate"] -Format 'yyyy-MM-dd')"
            }
            $Row["LatestUpdateType"] = $LatestUpdates.LatestUpdateType
        }
    }
}
#endregion

######################################################
## CALCULATE COMPLIANCE AGAINST RECENT WINDOWS CU'S ##
######################################################
#region CalculateCUCompiance
# Latest 
foreach ($Row in $DevicesDatatable)
{
    [array]$LatestUpdates = @($row["LatestRegularUpdate"],$row["LatestPreviewUpdate"],$row["LatestOutofBandUpdate"])
    If ($LatestUpdates.Count -ge 1)
    {
        [array]$LatestUpdatesArray = @()
        foreach ($LatestUpdate in $LatestUpdates)
        {
            If ($LatestUpdate.length -gt 0 -and $LatestUpdate -isnot [System.DBNull])
            {
                $LatestUpdatesArray += [System.Version]$LatestUpdate
            }
        }
        $THELatestUpdate = ($LatestUpdatesArray | Sort Minor -Descending | Select -First 1)
        If ([System.Version]$Row["CurrentPatchLevel"] -ge $THELatestUpdate)
        {
            $Row["ComplianceStateLatest"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStateLatest"] = "Out-of-date" 
        }
    }
    Else 
    {
        $Row["ComplianceStateLatest"] = "N/A"
    }
}

# Latest Regular (Monthly B)
foreach ($Row in $DevicesDatatable)
{
    If ($Row["LatestRegularUpdate"].Length -eq 0 -or $Row["LatestRegularUpdate"].GetType() -eq [System.DBNull])
    {
        $Row["ComplianceStateRegular"] = "N/A"
    }
    Else
    {
        If ([System.Version]$Row["CurrentPatchLevel"] -ge [System.Version]$Row["LatestRegularUpdate"])
        {
            $Row["ComplianceStateRegular"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStateRegular"] = "Out-of-date"
        }
    }
}

# Latest Preview
foreach ($Row in $DevicesDatatable)
{
    If ($Row["LatestPreviewUpdate"].Length -eq 0 -or $Row["LatestPreviewUpdate"].GetType() -eq [System.DBNull])
    {
        $Row["ComplianceStatePreview"] = "N/A"
    }
    Else
    {
        If ([System.Version]$Row["CurrentPatchLevel"] -ge [System.Version]$Row["LatestPreviewUpdate"])
        {
            $Row["ComplianceStatePreview"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStatePreview"] = "Out-of-date"
        }
    }
}

# Latest OOB
foreach ($Row in $DevicesDatatable)
{
    If ($Row["LatestOutofBandUpdate"].Length -eq 0 -or $Row["LatestOutofBandUpdate"].GetType() -eq [System.DBNull])
    {
        $Row["ComplianceStateOutofBand"] = "N/A"
    }
    Else
    {
        If ([System.Version]$Row["CurrentPatchLevel"] -ge [System.Version]$Row["LatestOutofBandUpdate"])
        {
            $Row["ComplianceStateOutofBand"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStateOutofBand"] = "Out-of-date"
        }
    }
}

# Latest Regular (Monthly B) Less 1
foreach ($Row in $DevicesDatatable)
{
    If ($Row["LatestRegularUpdateLess1"].Length -eq 0 -or $Row["LatestRegularUpdateLess1"].GetType() -eq [System.DBNull])
    {
        $Row["ComplianceStateRegularLess1"] = "N/A"
    }
    Else
    {
        If ([System.Version]$Row["CurrentPatchLevel"] -ge [System.Version]$Row["LatestRegularUpdateLess1"])
        {
            $Row["ComplianceStateRegularLess1"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStateRegularLess1"] = "Out-of-date"
        }
    }
}

# Latest Preview Less 1
foreach ($Row in $DevicesDatatable)
{
    If ($Row["LatestPreviewUpdateLess1"].Length -eq 0 -or $Row["LatestPreviewUpdateLess1"].GetType() -eq [System.DBNull])
    {
        $Row["ComplianceStatePreviewLess1"] = "N/A"
    }
    Else
    {
        If ([System.Version]$Row["CurrentPatchLevel"] -ge [System.Version]$Row["LatestPreviewUpdateLess1"])
        {
            $Row["ComplianceStatePreviewLess1"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStatePreviewLess1"] = "Out-of-date"
        }
    }
}

# Latest OOB Less 1
foreach ($Row in $DevicesDatatable)
{
    If ($Row["LatestOutofBandUpdateLess1"].Length -eq 0 -or $Row["LatestOutofBandUpdateLess1"].GetType() -eq [System.DBNull])
    {
        $Row["ComplianceStateOutofBandLess1"] = "N/A"
    }
    Else
    {
        If ([System.Version]$Row["CurrentPatchLevel"] -ge [System.Version]$Row["LatestOutofBandUpdateLess1"])
        {
            $Row["ComplianceStateOutofBandLess1"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStateOutofBandLess1"] = "Out-of-date"
        }
    }
}

# Latest Regular (Monthly B) Less 2
foreach ($Row in $DevicesDatatable)
{
    If ($Row["LatestRegularUpdateLess2"].Length -eq 0 -or $Row["LatestRegularUpdateLess2"].GetType() -eq [System.DBNull])
    {
        $Row["ComplianceStateRegularLess2"] = "N/A"
    }
    Else
    {
        If ([System.Version]$Row["CurrentPatchLevel"] -ge [System.Version]$Row["LatestRegularUpdateLess2"])
        {
            $Row["ComplianceStateRegularLess2"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStateRegularLess2"] = "Out-of-date"
        }
    }
}

# Latest Preview Less 2
foreach ($Row in $DevicesDatatable)
{
    If ($Row["LatestPreviewUpdateLess2"].Length -eq 0 -or $Row["LatestPreviewUpdateLess2"].GetType() -eq [System.DBNull])
    {
        $Row["ComplianceStatePreviewLess2"] = "N/A"
    }
    Else
    {
        If ([System.Version]$Row["CurrentPatchLevel"] -ge [System.Version]$Row["LatestPreviewUpdateLess2"])
        {
            $Row["ComplianceStatePreviewLess2"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStatePreviewLess2"] = "Out-of-date"
        }
    }
}

# Latest OOB Less 2
foreach ($Row in $DevicesDatatable)
{
    If ($Row["LatestOutofBandUpdateLess2"].Length -eq 0 -or $Row["LatestOutofBandUpdateLess2"].GetType() -eq [System.DBNull])
    {
        $Row["ComplianceStateOutofBandLess2"] = "N/A"
    }
    Else
    {
        If ([System.Version]$Row["CurrentPatchLevel"] -ge [System.Version]$Row["LatestOutofBandUpdateLess2"])
        {
            $Row["ComplianceStateOutofBandLess2"] = "Up-to-date"
        }
        else 
        {
            $Row["ComplianceStateOutofBandLess2"] = "Out-of-date"
        }
    }
}
#endregion

#################################
## CALCULATE OS SUPPORT STATUS ##
#################################
#region CalculateOSSupportStatus
$Query = "
SU_OSSupportMatrix_CL 
| top-nested 1 of TimeGenerated by temp=max(1),
    top-nested of Windows_Release=Windows_Release_s by temp1=max(1),
    top-nested of Version=Version_s by temp2=max(1),
    top-nested of StartDate=StartDate_s by temp3=max(1),
    top-nested of EndDate=EndDate_s by temp4=max(1),
    top-nested of SupportPeriodInDays=SupportPeriodInDays_d by temp5=max(1),
    top-nested of InSupport=InSupport_s by temp6=max(1),
    top-nested of SupportDaysRemaining=SupportDaysRemaining_d by temp7=max(1),
    top-nested of EditionFamily=EditionFamily_s by temp8=max(1)
| project-away temp*
| order by Windows_Release,Version,EditionFamily desc 
"
$Result = Invoke-AzOperationalInsightsQuery -Workspace $Workspace -Query $Query -Timespan (New-TimeSpan -Hours 24) -ErrorAction Stop
$iSupportMatrix = $Result.Results

foreach ($Row in $DevicesDatatable.Rows)
{
    If ($Row["ProductName"] -match "Enterprise" -or $Row["ProductName"] -match "Education")
    {
        $SupportInfo = $iSupportMatrix.Where({$_.Windows_Release -eq $Row["Windows_Release"] -and $_.Version -eq $Row["OSVersion"] -and $_.EditionFamily -eq "Enterprise, Education and IoT Enterprise"})
    }
    Else
    {
        $SupportInfo = $iSupportMatrix.Where({$_.Windows_Release -eq $Row["Windows_Release"] -and $_.Version -eq $Row["OSVersion"] -and $_.EditionFamily -eq "Home, Pro, Pro Education and Pro for Workstations"})
    }
    If ($SupportInfo.InSupport -eq "True" -and [int]$SupportInfo.SupportDaysRemaining -le 30)
    {
        $Row["OSVersionSupportStatus"] = "Support ending in 30 days or less"
    }
    ElseIf ($SupportInfo.InSupport -eq "True")
    {
        $Row["OSVersionSupportStatus"] = "In support"
    }
    ElseIf ($SupportInfo.InSupport -eq "False")
    {
        $Row["OSVersionSupportStatus"] = "Support ended"
    }
    $Row["SupportDaysRemaining"] = $SupportInfo.SupportDaysRemaining
}
#endregion

####################################################
## CALCULATE LATEST MONTHLY B INSTALLATION STATUS ##
####################################################
#region CalculateBStatus
foreach ($Row in $DevicesDatatable.Rows)
{
    If ($Row["QualityUpdatesDeferralInDays"].GetType() -eq [System.DBNull])
    {
        [int]$QualityUpdatesDeferralInDays = 0
    }
    else 
    {
        [int]$QualityUpdatesDeferralInDays = $Row["QualityUpdatesDeferralInDays"] 
    }

    $Row["LatestRegularUpdateStatus"] = "Missing"
    $DuckandRun = $false
    If ($Row["ComplianceStateRegular"] -eq "Up-to-date")
    {
        $Row["LatestRegularUpdateStatus"] = "Installed"
        $DuckandRun = $true
    }
    ElseIf ($Row["PatchRebootRequired"] -eq "TRUE")
    {
        $Row["LatestRegularUpdateStatus"] = "Installed pending restart"
        $DuckandRun = $true
    }
    ElseIf ($Row["QualityUpdatesPaused"] -eq "1")
    {
        $Row["LatestRegularUpdateStatus"] = "Quality updates paused"
        $DuckandRun = $true
    }

    If ($QualityUpdatesDeferralInDays -ge 1 -and $DuckandRun -eq $false)
    {
        $LatestRegularUpdateReleaseDate = $Row["LatestRegularUpdateReleaseDate"]
        If (-not [System.Convert]::IsDBNull($LatestRegularUpdateReleaseDate))
        {
            If ($LatestRegularUpdateReleaseDate -match "AM")
            {
                $LatestRegularUpdateReleaseDate = $LatestRegularUpdateReleaseDate.Replace("AM","").Trim()
            }            
            try 
            {
                $LatestRegularUpdateReleaseDate = $LatestRegularUpdateReleaseDate | Get-Date -ErrorAction Stop   
            }
            catch 
            {
                Write-Warning "Failed to convert $LatestRegularUpdateReleaseDate to DateTime"
            }
            If ($LatestRegularUpdateReleaseDate -is [datetime])
            {
                If (([DateTime]::UtcNow - $LatestRegularUpdateReleaseDate).TotalDays -lt $QualityUpdatesDeferralInDays)
                {
                    $Row["LatestRegularUpdateStatus"] = "Deferred $QualityUpdatesDeferralInDays days"
                    $DuckandRun = $true
                }
            }
        }  
    }
    # Language arrays
    $InstallationArray = @(
        'Installation' # English,German
        'Installazione' # Italian
    )
    $FailureArray = @(
        'Failure' # English
        'Fehler' # German
        'Errore' # Italian
    )
    $StartedArray = @(
        'Started' # English
        'Gestartet' # German
        'Avviato' # Italian
    )
    If ($DuckandRun -eq $false -and $Row["UpdateActivity"] -in $InstallationArray -and $Row["UpdateStatus"] -in $FailureArray)
    {
        $Row["LatestRegularUpdateStatus"] = "Installation failure"
    }
    ElseIf ($DuckandRun -eq $false -and $Row["UpdateActivity"] -in $InstallationArray -and $Row["UpdateStatus"] -in $StartedArray)
    {
        $Row["LatestRegularUpdateStatus"] = "Installation started"
    }  
    ElseIf ($DuckandRun -eq $false -and $Row["LatestRegularUpdateName"].Length -eq 0)
    {
        $Row["LatestRegularUpdateStatus"] = "N/A"
    }
    ElseIf ($DuckandRun -eq $false -and $Row["ComplianceStateRegular"] -eq "Out-of-date")
    {
        $Row["LatestRegularUpdateStatus"] = "Missing"
    }
}
#endregion

#########################################
## POST THE SUMMARISED COMPLIANCE DATA ##
#########################################
#region PostData
$SummarizationTime = Get-Date ([DateTime]::UtcNow) -Format "s"
foreach ($Row in $DevicesDatatable.Rows)
{
    $Row["SummarizationTime"] = $SummarizationTime
}
$Json = $DevicesDatatable.Rows | Select $DevicesDatatable.Columns.ColumnName | ConvertTo-Json -Compress
[int]$JsonSize = [System.Text.Encoding]::UTF8.GetByteCount($Json) / 1MB
# If the resulting JSON is larger than the posting limit (30MB) we split in two
If ($JsonSize -gt 30)
{
    $RowCount = $DevicesDatatable.Rows.Count
    $FirstHalfCount = [math]::Floor($rowCount / 2)
    $Json1 = $DevicesDatatable.Rows | Select -First $FirstHalfCount -Property $DevicesDatatable.Columns.ColumnName | ConvertTo-Json -Compress
    $Json2 = $DevicesDatatable.Rows | Select -Skip $FirstHalfCount -Property $DevicesDatatable.Columns.ColumnName | ConvertTo-Json -Compress

    $Json1,$Json2 | foreach {
        $Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($_)) -logType "SU_ClientComplianceStatus"
        If ($Result.GetType().Name -eq "ErrorRecord")
        {
            Write-Error -Exception $Result.Exception
        }
        else 
        {
            $Result.StatusCode  
        }
    }
}
else 
{
    $Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType "SU_ClientComplianceStatus"
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
