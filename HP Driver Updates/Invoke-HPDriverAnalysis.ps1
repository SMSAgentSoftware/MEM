###############################
## HP DRIVER UPDATE ANALYZER ##
###############################


################
## Parameters ##
################
# WorkspaceID of the Log Analytics workspace
$script:WorkspaceID = "<YourWorkspaceID>" 
# Primary Key of the Log Analytics workspace
$script:PrimaryKey = "<YourPrimaryKey>" 
# The name of the table to create/use in the Log Analytics workspace
$script:LogName = "HPDriverUpdates" 
# The name of the parent folder and registry key that we'll work with, eg your company or IT dept name
$ParentFolderName = "Contoso" 
#  The name of the child folder and registry key that we'll work with
$ChildFolderName = "HP_Driver_Analysis" 
# The minimum number of hours in between each run of this script
[int]$MinimumFrequency = 5
# Static web page of the HP Image Assistant
$HPIAWebUrl = "https://ftp.hp.com/pub/caps-softpaq/cmit/HPIA.html" 
# Set the security protocol. Must include Tls1.2.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12,[Net.SecurityProtocolType]::Tls13
# to speed up web requests
$ProgressPreference = 'SilentlyContinue' 


###############
## Functions ##
###############
#region 
# Function write to a log file in ccmtrace format
Function script:Write-Log {

    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
		
        [Parameter()]
        [ValidateSet(1, 2, 3)] # 1-Info, 2-Warning, 3-Error
        [int]$LogLevel = 1,

        [Parameter(Mandatory = $true)]
        [string]$Component,

        [Parameter(Mandatory = $false)]
        [object]$Exception
    )
   
    If ($Exception)
    {
        [String]$Message = "$Message" + "$Exception"
    }

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), $Component, $LogLevel
    $Line = $Line -f $LineFormat
    
    # Write to log
    Add-Content -Value $Line -Path $LogFile -ErrorAction SilentlyContinue

}

# Function to do a log rollover
Function Rollover-Log {
   
    # Create the log file
    If (!(Test-Path $LogFile))
	{
	    $null = New-Item $LogFile -Type File
	}
    
    # Log rollover
    If ([math]::Round((Get-Item $LogFile).Length / 1KB) -gt 2000)
    {
        Write-Log "Log has reached 2MB. Rolling over..."
        Rename-Item -Path $LogFile -NewName "HP_Driver_Analysis-$(Get-Date -Format "yyyyMMdd-hhmmss").log"
        $null = New-Item $LogFile -Type File
    } 

    # Remove oldest log
    If ((Get-ChildItem $ParentDirectory -Name "HP_Driver_Analysis*.log").Count -eq 3)
    {
        (Get-ChildItem -Path $ParentDirectory -Filter "HP_Driver_Analysis*.log" | 
            select FullName,LastWriteTime | 
            Sort LastWriteTime | 
            Select -First 1).FullName | Remove-Item  
    }
		
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

    try {
        $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing
    }
    catch {
        $response = $_#.Exception.Response
    }
    
    return $response
}
#endregion


#################
## Preparation ##
#################
#region
# Custom class and list
class Recommendation {
    [string]$TargetComponent
    [string]$TargetVersion
    [string]$ReferenceVersion
    [string]$Comments
    [string]$SoftPaqId
    [string]$Name
    [string]$Type
    [string]$Model
    [string]$IntuneDeviceID
    [string]$AADDeviceID
    [string]$ComputerName
    [string]$InventoryDate
}
$Recommendations = [System.Collections.Generic.List[Recommendation]]::new()

# Safeguard to prevent execution on non-HP workstations
$Manufacturer = Get-CimInstance -ClassName Win32_ComputerSystem -Property Manufacturer -ErrorAction SilentlyContinue | Select -ExpandProperty Manufacturer
If ($Manufacturer -notin ('HP','Hewlett-Packard'))
{
    Write-Output "Not an HP workstation"
    Return
}
#endregion


##########################
## Create Registry Keys ##
##########################
#region
$RegRoot = "HKLM:\SOFTWARE"
$FullRegPath = "$RegRoot\$ParentFolderName\$ChildFolderName"
If (!(Test-Path $RegRoot\$ParentFolderName))
{
    $null = New-Item -Path $RegRoot -Name $ParentFolderName -Force
}
If (!(Test-Path $FullRegPath))
{
    $null = New-Item -Path $RegRoot\$ParentFolderName -Name $ChildFolderName -Force
}
#endregion


#############################
## Check the run frequency ##
#############################
#region
# This is to ensure that the script does not attempt to run more frequently than the defined schedule in the MinimumFrequency value
$LatestRunStartTime = Get-ItemProperty -Path $FullRegPath -Name LatestRunStartTime -ErrorAction SilentlyContinue | Select -ExpandProperty LatestRunStartTime | Get-Date -ErrorAction SilentlyContinue
if ($null -ne $LatestRunStartTime)
{
    If (((Get-Date) - $LatestRunStartTime).TotalHours -le $MinimumFrequency)
    {
        Write-Output "Minimum threshold for script re-run has not yet been met"
        Return
    }
}
Set-ItemProperty -Path $FullRegPath -Name LatestRunStartTime -Value (Get-Date -Format "s") -Force
#endregion


################################################################################
## Check if an inventory is already running to avoid simultaneous executions  ##
################################################################################
#region
# This is necessary due to the fact that proactive remediations can run in multiple contexts
$ExecutionStatus = Get-ItemProperty -Path $FullRegPath -Name ExecutionStatus -ErrorAction SilentlyContinue | Select -ExpandProperty ExecutionStatus | Get-Date -ErrorAction SilentlyContinue
If ($ExecutionStatus -eq "Running")
{
    Write-Output "Another execution is currently running"
    Return
}
else 
{
    Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Running" -Force
}
#endregion


################################
## Create Directory Structure ##
################################
#region
$RootFolder = $env:ProgramData
$ChildFolderName2 = Get-Date -Format "yyyy-MMM-dd_HH.mm.ss"
$script:ParentDirectory = "$RootFolder\$ParentFolderName\$ChildFolderName"
$script:WorkingDirectory = "$RootFolder\$ParentFolderName\$ChildFolderName\$ChildFolderName2"
try 
{
    [void][System.IO.Directory]::CreateDirectory($WorkingDirectory)
}
catch 
{
    throw $_.Exception.Message
}
$script:LogFile = "$ParentDirectory\HP_Driver_Analysis.log"
Rollover-Log
Write-Log -Message "#########################" -Component "Preparation"
Write-Log -Message "## Starting HP Driver Analysis ##" -Component "Preparation"
Write-Log -Message "#########################" -Component "Preparation"
#endregion


#################################
## Disable IE First Run Wizard ##
#################################
#region
# This prevents an error running Invoke-WebRequest when IE has not yet been run in the current context
Write-Log -Message "Disabling IE first run wizard" -Component "Preparation"
$null = New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft" -Name "Internet Explorer" -Force
$null = New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer" -Name "Main" -Force
$null = New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -PropertyType DWORD -Value 1 -Force
#endregion


##########################
## Get latest HPIA Info ##
##########################
#region
Write-Log -Message "Finding info for latest version of HP Image Assistant (HPIA)" -Component "DownloadHPIA"
try
{
    $HTML = Invoke-WebRequest -Uri $HPIAWebUrl -ErrorAction Stop
}
catch 
{
    Write-Log -Message "Failed to download the HPIA web page. $($_.Exception.Message)" -Component "DownloadHPIA" -LogLevel 3
    Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
    Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    throw "Failed to download the HPIA web page. $($_.Exception.Message)"
}
$HPIASoftPaqNumber = ($HTML.Links | Where {$_.href -match "hp-hpia-"}).outerText
$HPIADownloadURL = ($HTML.Links | Where {$_.href -match "hp-hpia-"}).href
$HPIAFileName = $HPIADownloadURL.Split('/')[-1]
Write-Log -Message "HPIA SoftPaq number is $HPIASoftPaqNumber" -Component "DownloadHPIA"
Write-Log -Message "HPIA download URL is $HPIADownloadURL" -Component "DownloadHPIA"
#endregion


##########################################################
## Check if the current HPIA version is already present ##
##########################################################
#region
$File = Get-Item -Path $ParentDirectory\HPIA\HPImageAssistant.exe -ErrorAction SilentlyContinue
If ($null -eq $File)
{
    Write-Log -Message "HP Image Assistant not found locally. Proceed with download" -Component "DownloadHPIA"
    $DownloadHPIA = $true
}
else 
{
    $FileVersion = $HPIAFileName.Split('-')[-1].TrimEnd('.exe')
    $ProductVersion = $File.VersionInfo.ProductVersion
    If ($ProductVersion -match $FileVersion)
    {
        Write-Log -Message "HP Image Assistant was found locally at the current version. No need to download" -Component "DownloadHPIA"
    }
    else 
    {
        Write-Log -Message "HP Image Assistant was found locally but not at the current version. Proceed to download" -Component "DownloadHPIA"
        $DownloadHPIA = $true
    }
}
#endregion


###################
## Download HPIA ##
###################
#region
If ($DownloadHPIA -eq $true)
{
    Write-Log -Message "Downloading the HPIA" -Component "DownloadHPIA"
    try 
    {
        $ExistingBitsJob = Get-BitsTransfer -Name "$HPIAFileName" -AllUsers -ErrorAction SilentlyContinue
        If ($ExistingBitsJob)
        {
            Write-Log -Message "An existing BITS tranfer was found. Cleaning it up." -Component "DownloadHPIA" -LogLevel 2
            Remove-BitsTransfer -BitsJob $ExistingBitsJob
        }
        $BitsJob = Start-BitsTransfer -Source $HPIADownloadURL -Destination $ParentDirectory\$HPIAFileName -Asynchronous -DisplayName "$HPIAFileName" -Description "HPIA download" -RetryInterval 60 -ErrorAction Stop 
        do {
            Start-Sleep -Seconds 5
            $Progress = [Math]::Round((100 * ($BitsJob.BytesTransferred / $BitsJob.BytesTotal)),2)
            Write-Log -Message "Downloaded $Progress`%" -Component "DownloadHPIA"
        } until ($BitsJob.JobState -in ("Transferred","Error"))
        If ($BitsJob.JobState -eq "Error")
        {
            Write-Log -Message "BITS tranfer failed: $($BitsJob.ErrorDescription)" -Component "DownloadHPIA" -LogLevel 3
            Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
            Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
            throw "BITS tranfer failed: $($BitsJob.ErrorDescription)"
        }
        Write-Log -Message "Download is finished" -Component "DownloadHPIA"
        Complete-BitsTransfer -BitsJob $BitsJob
        Write-Log -Message "BITS transfer is complete" -Component "DownloadHPIA"
    }
    catch 
    {
        Write-Log -Message "Failed to start a BITS transfer for the HPIA: $($_.Exception.Message)" -Component "DownloadHPIA" -LogLevel 3
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        throw "Failed to start a BITS transfer for the HPIA: $($_.Exception.Message)" 
    }
}
#endregion


##################
## Extract HPIA ##
##################
#region
If ($DownloadHPIA -eq $true)
{
    Write-Log -Message "Extracting the HPIA" -Component "Analyze"
    try 
    {
        $Process = Start-Process -FilePath $ParentDirectory\$HPIAFileName -WorkingDirectory $ParentDirectory -ArgumentList "/s /f .\HPIA\ /e" -NoNewWindow -PassThru -Wait -ErrorAction Stop
        Start-Sleep -Seconds 5
        If (Test-Path $ParentDirectory\HPIA\HPImageAssistant.exe)
        {
            Write-Log -Message "Extraction complete" -Component "Analyze"
        }
        Else  
        {
            Write-Log -Message "HPImageAssistant not found!" -Component "Analyze" -LogLevel 3
            Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
            Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
            throw "HPImageAssistant not found!"
        }
    }
    catch 
    {
        Write-Log -Message "Failed to extract the HPIA: $($_.Exception.Message)" -Component "Analyze" -LogLevel 3
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        throw "Failed to extract the HPIA: $($_.Exception.Message)"
    }
}
#endregion


#########################################
## Analyze available updates with HPIA ##
#########################################
#region
Write-Log -Message "Analyzing system for available updates" -Component "Analyze"
try 
{
    $Process = Start-Process -FilePath $ParentDirectory\HPIA\HPImageAssistant.exe -WorkingDirectory $ParentDirectory -ArgumentList "/Operation:Analyze /Category:ALL /Selection:All /Action:List /Silent /ReportFolder:$WorkingDirectory\Report" -NoNewWindow -PassThru -Wait -ErrorAction Stop
    If ($Process.ExitCode -eq 0)
    {
        Write-Log -Message "Analysis complete" -Component "Analyze"
    }
    elseif ($Process.ExitCode -eq 256) 
    {
        Write-Log -Message "The analysis returned no recommendation. No updates are available at this time" -Component "Analyze" -LogLevel 2
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Complete" -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        Return
    }
    elseif ($Process.ExitCode -eq 4096) 
    {
        Write-Log -Message "This platform is not supported!" -Component "Analyze" -LogLevel 2
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Not supported" -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        Return
    }
    Else
    {
        Write-Log -Message "Process exited with code $($Process.ExitCode). Expecting 0." -Component "Analyze" -LogLevel 3
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        throw "Process exited with code $($Process.ExitCode). Expecting 0."
    }
}
catch 
{
    Write-Log -Message "Failed to start the HPImageAssistant.exe: $($_.Exception.Message)" -Component "Analyze" -LogLevel 3
    Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
    Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    throw "Failed to start the HPImageAssistant.exe: $($_.Exception.Message)"
}
#endregion


###########################
## Review the XML report ##
###########################
#region
Write-Log -Message "Reading xml report" -Component "Analyze"
try 
{
    $XMLFile = Get-ChildItem -Path "$WorkingDirectory\Report" -Recurse -Include *.xml -ErrorAction Stop
    If ($XMLFile)
    {
        try 
        {
            [xml]$XML = Get-Content -Path $XMLFile.FullName -ErrorAction Stop
            [array]$SoftwareRecommendations = $xml.HPIA.Recommendations.Software.Recommendation
            [array]$DriverRecommendations = $xml.HPIA.Recommendations.Drivers.Recommendation
            [array]$BIOSRecommendations = $xml.HPIA.Recommendations.BIOS.Recommendation
            [array]$FirmwareRecommendations = $xml.HPIA.Recommendations.Firmware.Recommendation

            If ($SoftwareRecommendations.Count -ge 1)
            {
                Write-Log -Message "Found $($SoftwareRecommendations.Count) software recommendations" -Component "Analyze"
                foreach ($Item in $SoftwareRecommendations)
                {
                    $Recommendation = [Recommendation]::new()
                    $Recommendation.TargetComponent = $item.TargetComponent
                    $Recommendation.TargetVersion = $item.TargetVersion
                    $Recommendation.ReferenceVersion = $item.ReferenceVersion
                    $Recommendation.Comments = $item.Comments
                    $Recommendation.SoftPaqId = $item.Solution.Softpaq.Id
                    $Recommendation.Name = $item.Solution.Softpaq.Name
                    $Recommendation.Type = "Software"
                    $Recommendations.Add($Recommendation)
                    Write-Log -Message ">> $($Recommendation.SoftPaqId): $($Recommendation.Name) ($($Recommendation.ReferenceVersion))" -Component "Analyze"
                }
            }

            If ($DriverRecommendations.Count -ge 1)
            {
                Write-Log -Message "Found $($DriverRecommendations.Count) driver recommendations" -Component "Analyze"
                foreach ($Item in $DriverRecommendations)
                {
                    $Recommendation = [Recommendation]::new()
                    $Recommendation.TargetComponent = $item.TargetComponent
                    $Recommendation.TargetVersion = $item.TargetVersion
                    $Recommendation.ReferenceVersion = $item.ReferenceVersion
                    $Recommendation.Comments = $item.Comments
                    $Recommendation.SoftPaqId = $item.Solution.Softpaq.Id
                    $Recommendation.Name = $item.Solution.Softpaq.Name
                    $Recommendation.Type = "Driver"
                    $Recommendations.Add($Recommendation)
                    Write-Log -Message ">> $($Recommendation.SoftPaqId): $($Recommendation.Name) ($($Recommendation.ReferenceVersion))" -Component "Analyze"
                }
            }

            If ($BIOSRecommendations.Count -ge 1)
            {
                Write-Log -Message "Found $($BIOSRecommendations.Count) BIOS recommendations" -Component "Analyze"
                foreach ($Item in $BIOSRecommendations)
                {
                    $Recommendation = [Recommendation]::new()
                    $Recommendation.TargetComponent = $item.TargetComponent
                    $Recommendation.TargetVersion = $item.TargetVersion
                    $Recommendation.ReferenceVersion = $item.ReferenceVersion
                    $Recommendation.Comments = $item.Comments
                    $Recommendation.SoftPaqId = $item.Solution.Softpaq.Id
                    $Recommendation.Name = $item.Solution.Softpaq.Name
                    $Recommendation.Type = "BIOS"
                    $Recommendations.Add($Recommendation)
                    Write-Log -Message ">> $($Recommendation.SoftPaqId): $($Recommendation.Name) ($($Recommendation.ReferenceVersion))" -Component "Analyze"
                }
            }

            If ($FirmwareRecommendations.Count -ge 1)
            {
                Write-Log -Message "Found $($FirmwareRecommendations.Count) firmware recommendations" -Component "Analyze"
                foreach ($Item in $FirmwareRecommendations)
                {
                    $Recommendation = [Recommendation]::new()
                    $Recommendation.TargetComponent = $item.TargetComponent
                    $Recommendation.TargetVersion = $item.TargetVersion
                    $Recommendation.ReferenceVersion = $item.ReferenceVersion
                    $Recommendation.Comments = $item.Comments
                    $Recommendation.SoftPaqId = $item.Solution.Softpaq.Id
                    $Recommendation.Name = $item.Solution.Softpaq.Name
                    $Recommendation.Type = "Firmware"
                    $Recommendations.Add($Recommendation)
                    Write-Log -Message ">> $($Recommendation.SoftPaqId): $($Recommendation.Name) ($($Recommendation.ReferenceVersion))" -Component "Analyze"
                }
            }

            If ($DriverRecommendations.Count -eq 0 -and $SoftwareRecommendations.Count -eq 0 -and $BIOSRecommendations.Count -eq 0 -and $FirmwareRecommendations.Count -eq 0)
            {
                Write-Log -Message "No recommendations found at this time" -Component "Analyze"
                Write-Log -Message "This driver analysis is complete. Have a nice day!" -Component "Completion"
                Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Complete" -Force
                Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
                Return
            }
        }
        catch 
        {
            Write-Log -Message "Failed to parse the XML file: $($_.Exception.Message)" -Component "Analyze" -LogLevel 3
            Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
            Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
            throw "Failed to parse the XML file: $($_.Exception.Message)" 
        }
    }
    Else  
    {
        Write-Log -Message "Failed to find an XML report." -Component "Analyze" -LogLevel 3
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        throw "Failed to find an XML report."
    }
}
catch 
{
    Write-Log -Message "Failed to find an XML report: $($_.Exception.Message)" -Component "Analyze" -LogLevel 3
    Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
    Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    throw "Failed to find an XML report: $($_.Exception.Message)"
}
#endregion


##########################
## GET INTUNE DEVICE ID ##
##########################
#region
$IntuneCert = (Get-ChildItem Cert:\*\MY -Recurse | Where {$_.Issuer -eq "CN=Microsoft Intune MDM Device CA"})
If ($null -ne $IntuneCert)
{
    # Sometimes an expired cert may still exist
    if ($IntuneCert.GetType().BaseType.Name -eq "Array")
    {
        $IntuneCert = $IntuneCert | Sort NotAfter -Descending | Select -First 1 
    }
    $IntuneDeviceID = $IntuneCert.Subject.Replace('CN=','')
}
else 
{
    Write-Log "No Intune device Id could be found for this device" -Component "Post" -LogLevel 2
}
#endregion


#######################
## GET AAD DEVICE ID ##
#######################
#region
$AADCert = (Get-ChildItem Cert:\Localmachine\MY | Where {$_.Issuer -match "CN=MS-Organization-Access"})
If ($null -ne $AADCert)
{
    $AADDeviceID = $AADCert.Subject.Replace('CN=','')
}
else 
{
    Write-Log "No Azure AD device Id could be found for this device" -Component "Post" -LogLevel 2
}
#endregion


###################################
## POST THE DATA TO LA WORKSPACE ##
###################################
#region
$Model = Get-CimInstance -ClassName Win32_ComputerSystem -Property Model -ErrorAction SilentlyContinue | Select -ExpandProperty Model
$InventoryDate = Get-Date ([DateTime]::UtcNow) -Format "s"
foreach ($item in $Recommendations)
{
    $item.IntuneDeviceID = $IntuneDeviceID
    $item.AADDeviceID = $AADDeviceID
    $item.ComputerName = $env:COMPUTERNAME 
    $item.InventoryDate = $InventoryDate
    $item.Model = $Model
}

# First test/wait for access to the workspace
$Endpoint = "$WorkspaceID.ods.opinsights.azure.com"
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$i = 0
do {
    If ($i -gt 0)
    {
        Start-Sleep -Seconds 10
    }
    try 
    {
        $tcp = [System.Net.Sockets.TcpClient]::new($Endpoint, 443)           
    }
    catch 
    {
        Write-Log -Message "Unable to connect to the Log Analytics endpoint: $($_.Exception.InnerException.Message). Retry logic is active." -Component "Post" -LogLevel 2
        $i ++
    }
}   
Until ($tcp.Connected -eq $true -or $Stopwatch.Elapsed.TotalSeconds -ge 320)
$Stopwatch.Stop() 

If ($Stopwatch.Elapsed.TotalSeconds -ge 320)
{
    Write-Log -Message "Gave up trying to connect to the Log Analytics endpoint. The log entry will not be posted." -Component "Post" -LogLevel 3
    Return
}

$tcp.Close()
$tcp.Dispose()
$Json = $Recommendations | ConvertTo-Json -Compress
$Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($Json)) -logType $LogName
If ($Result.GetType().Name -eq "WebResponseObject")
{
    If ($Result.StatusCode -eq 200)
    {
        Write-Log "Successfully posted to log analytics workspace" -Component "Post" 
    }
    else 
    {
        $UnsuccessfullySent ++
        If ($null -ne $Result.Response)
        {
            Write-Log "Failed to post to log analytics workspace! $($Result.Response.StatusCode) | $($Result.Response.StatusDescription)" -Component "Post" -LogLevel 3  
        }
        Else 
        {
            Write-Log "Failed to post to log analytics workspace! $($Result.Message)" -Component "Post"  -LogLevel 3  
        }
    }
}
ElseIf ($Result.GetType().Name -eq "ErrorRecord")
{
    Write-Log "Failed to post to log analytics workspace! $($Result.Exception.Message)" -Component "Post"  -LogLevel 3  
}
Else
{
    Write-Log "Failed to post to log analytics workspace! $Result" -Component "Post"  -LogLevel 3  
}
#endregion


###############
## Finish up ##
###############
#region
Write-Log -Message "This driver analysis is complete. Have a nice day!" -Component "Completion"
Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Complete" -Force
Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
#endregion
