#######################
## HP DRIVER UPDATER ##
#######################


################
## Parameters ##
################
# Specify the update run type, one of :
#  "All" (for all available updates), 
#  "CriticalOnly" (for all updates marked as critical by HP),
#  "SpecificCategories" (for all available updates in a specific category or categories - enter the category name/s below)
#  "SpecificUpdates" (for a specific SoftPaq or SoftPaqs - enter the SoftPaqNumbers below)
$UpdateRunType = "CriticalOnly"
# Category names if applicable. Choose from Firmware, Driver, Software.
$Categories = @("Software","Driver")
# The SoftPaq numbers for the drivers to install if applicable.
$SoftPaqNumbers = @("sp143082") 
# WorkspaceID of the Log Analytics workspace
$script:WorkspaceID = "<YourWorkspaceID>" 
# Primary Key of the Log Analytics workspace
$script:PrimaryKey = "<YourPrimaryKey>" 
# The name of the table to create/use in the Log Analytics workspace
$script:LogName = "HPDriverUpdatesInstallLog" 
# The name of the parent folder and registry key that we'll work with, eg your company or IT dept name
$ParentFolderName = "Contoso" 
#  The name of the child folder and registry key that we'll work with
$ChildFolderName = "HP_Driver_Updates" 
# Maximum number of minutes to wait for an update to install
[int]$InstallTimeout = 5
# The minimum number of hours in between each run of this script. Used to minimize the possibility of multiple executions in different contexts.
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
        Rename-Item -Path $LogFile -NewName "HP_Driver_Updates-$(Get-Date -Format "yyyyMMdd-hhmmss").log"
        $null = New-Item $LogFile -Type File
    } 

    # Remove oldest log
    If ((Get-ChildItem $ParentDirectory -Name "HP_Driver_Updates*.log").Count -eq 3)
    {
        (Get-ChildItem -Path $ParentDirectory -Filter "HP_Driver_Updates*.log" | 
            select FullName,LastWriteTime | 
            Sort LastWriteTime | 
            Select -First 1).FullName | Remove-Item  
    }
		
}

# Create the function to create the authorization signature
# ref https://docs.microsoft.com/en-us/azure/azure-monitor/logs/data-collector-api
Function script:Build-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource)
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
Function script:Post-LogAnalyticsData($customerId, $sharedKey, $body, $logType)
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

# Function to post a log entry to log analytics
Function Post-LogEntry {
    Param([Object]$EntryInfo)
    # Get the device model  
    $Model = Get-CimInstance -ClassName Win32_ComputerSystem -Property Model -ErrorAction SilentlyContinue | Select -ExpandProperty Model
    # Get Intune device Id
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
    # Get AAD Device Id
    $AADCert = (Get-ChildItem Cert:\Localmachine\MY | Where {$_.Issuer -match "CN=MS-Organization-Access"})
    If ($null -ne $AADCert)
    {
        $AADDeviceID = $AADCert.Subject.Replace('CN=','')
    }
    else 
    {
        Write-Log "No Azure AD device Id could be found for this device" -Component "Post" -LogLevel 2
    }
    
    # If a log entry was created prior to download/install
    If ($EntryInfo -is [LogEntry])
    {
        $LogEntry = $EntryInfo
        $LogEntry.InstallDate = Get-Date ([DateTime]::UtcNow) -Format "s"
        $LogEntry.InstallStatus = $EntryInfo.InstallStatus.ToUpper()
        $LogEntry.Message = $EntryInfo.Message
        $LogEntry.SoftPaqNumber = $EntryInfo.SoftPaqNumber
    }
    else 
    {
        $LogEntry = [LogEntry]::new()
        $LogEntry.ReturnCode = $EntryInfo.ReturnCode   
        $LogEntry.RebootRequired = $EntryInfo.RebootRequired
        $LogEntry.Message = $EntryInfo.Message
        $LogEntry.InstallStatus = $EntryInfo.InstallStatus.ToUpper()
        $LogEntry.UpdateName = $EntryInfo.UpdateName
        $LogEntry.Version = $EntryInfo.Version
        $LogEntry.VendorVersion = $EntryInfo.VendorVersion
        $LogEntry.InstallDate = $EntryInfo.InstallDate
        $LogEntry.SoftPaqNumber = $EntryInfo.SoftPaqNumber
        $LogEntry.Category = $EntryInfo.Category
    }
    
    # Some common values for each log entry
    $LogEntry.AADDeviceID = $AADDeviceID
    $LogEntry.IntuneDeviceID = $IntuneDeviceID
    $LogEntry.ComputerName = $env:COMPUTERNAME
    $LogEntry.Model = $Model 

    # Test access to the endpoint here if prior to download/install
    If ($EntryInfo -is [LogEntry])
    {
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
        $tcp.Close()
        $tcp.Dispose()

        If ($Stopwatch.Elapsed.TotalSeconds -ge 320)
        {
            Write-Log -Message "Gave up trying to connect to the Log Analytics endpoint. The log entry will not be posted." -Component "Post" -LogLevel3
            Return
        }
    }

    # Post the data and handle the response
    $Json = $LogEntry | ConvertTo-Json -Compress
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
}
#endregion


#################
## Preparation ##
#################
#region
# Custom classes
class Recommendation {
    [string]$TargetComponent
    [string]$TargetVersion
    [string]$ReferenceVersion
    [string]$Comments
    [string]$SoftPaqId
    [string]$Name
    [string]$Type  
    [string]$CVAUrl
    [string]$ExeUrl
}
class ReturnCode {
    [string]$Code
    [string]$Status
    [string]$RebootRequired
    [string]$Message
}
class LogEntry {
    [string]$UpdateName
    [string]$Version
    [string]$VendorVersion
    [string]$InstallStatus
    [string]$ReturnCode
    [string]$RebootRequired
    [string]$Message
    [string]$SoftPaqNumber
    [string]$IntuneDeviceID
    [string]$AADDeviceID
    [string]$ComputerName
    [string]$Model
    [string]$InstallDate
    [string]$Category
}
class UpdateEntry {
    [string]$UpdateName
    [string]$Version
    [string]$VendorVersion
    [string]$InstallStatus
    [string]$ReturnCode
    [string]$RebootRequired
    [string]$Message
    [string]$SoftPaqNumber
    [string]$InstallDate
    [string]$ExeDownloadURL
    [string]$CVADownloadURL
    [string]$ExeFilename
    [string]$CVAFilename
    [string]$SHA256Hash
    [string]$SilentInstallCmd
    [string]$Category
    [System.Collections.Generic.List[ReturnCode]]$ReturnCodeList
}
$script:LogEntry = [LogEntry]::new()
$Recommendations = [System.Collections.Generic.List[Recommendation]]::new()
$UpdateEntries = [System.Collections.Generic.List[UpdateEntry]]::new()

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
$ParentRegPath = "$RegRoot\$ParentFolderName\$ChildFolderName"
$CreateRegPath = "SOFTWARE\$ParentFolderName\$ChildFolderName"
If (!(Test-Path $ParentRegPath))
{
    [void][Microsoft.Win32.Registry]::LocalMachine.CreateSubKey($CreateRegPath)
}
#endregion


#############################
## Check the run frequency ##
#############################
#region
# This is to ensure that the script does not attempt to run more frequently than the defined schedule in the MinimumFrequency value
$LatestRunStartTime = Get-ItemProperty -Path $ParentRegPath -Name LatestRunStartTime -ErrorAction SilentlyContinue | Select -ExpandProperty LatestRunStartTime | Get-Date -ErrorAction SilentlyContinue
if ($null -ne $LatestRunStartTime)
{
    If (((Get-Date) - $LatestRunStartTime).TotalHours -le $MinimumFrequency)
    {
        Write-Output "Minimum threshold for script re-run has not yet been met"
        Return
    }
}
Set-ItemProperty -Path $ParentRegPath -Name LatestRunStartTime -Value (Get-Date -Format "s") -Force
#endregion


################################################################################
## Check if an update run is already active to avoid simultaneous executions  ##
################################################################################
#region
# This is necessary due to the fact that proactive remediations can run in multiple contexts
$ExecutionStatus = Get-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -ErrorAction SilentlyContinue | Select -ExpandProperty ExecutionStatus | Get-Date -ErrorAction SilentlyContinue
If ($ExecutionStatus -eq "Running")
{
    Write-Output "Another execution is currently running"
    Return
}
else 
{
    Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Running" -Force
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
    $LogEntry.InstallStatus = "Not started"
    $LogEntry.Message = "Failed to create working directory"
    Post-LogEntry -EntryInfo $LogEntry
    Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Not started" -Force
    Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
    return
}
$script:LogFile = "$ParentDirectory\HP_Driver_Updates.log"
Rollover-Log
Write-Log -Message "###########################" -Component "Preparation"
Write-Log -Message "## Starting HP Driver Update Run ##" -Component "Preparation"
Write-Log -Message "###########################" -Component "Preparation"
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
    $LogEntry.InstallStatus = "Not started"
    $LogEntry.Message = "Failed to download the HPIA web page. $($_.Exception.Message)"
    Post-LogEntry -EntryInfo $LogEntry
    Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Not started" -Force
    Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
    Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    return
}
$HPIASoftPaqNumber = ($HTML.Links | Where {$_.href -match "hp-hpia-"}).outerText
$HPIADownloadURL = ($HTML.Links | Where {$_.href -match "hp-hpia-"}).href
$HPIAFileName = $HPIADownloadURL.Split('/')[-1]
Write-Log -Message "HPIA SoftPaq number is $HPIASoftPaqNumber" -Component "DownloadHPIA"
Write-Log -Message "Download URL is $HPIADownloadURL" -Component "DownloadHPIA"
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
            $LogEntry.InstallStatus = "Failed"
            $LogEntry.Message = "BITS tranfer failed: $($BitsJob.ErrorDescription)"
            Post-LogEntry -EntryInfo $LogEntry
            Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Failed" -Force
            Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
            Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
        Write-Log -Message "Download is finished" -Component "DownloadHPIA"
        Complete-BitsTransfer -BitsJob $BitsJob
        Write-Log -Message "BITS transfer is complete" -Component "DownloadHPIA"
    }
    catch 
    {
        Write-Log -Message "Failed to start a BITS transfer for the HPIA: $($_.Exception.Message)" -Component "DownloadHPIA" -LogLevel 3
        $LogEntry.InstallStatus = "Failed"
        $LogEntry.Message = "Failed to start a BITS transfer for the HPIA: $($_.Exception.Message)"
        Post-LogEntry -EntryInfo $LogEntry
        Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Failed" -Force
        Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        return
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
            $LogEntry.InstallStatus = "Failed"
            $LogEntry.Message = "HPImageAssistant not found!"
            Post-LogEntry -EntryInfo $LogEntry
            Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Failed" -Force
            Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
            Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
    }
    catch 
    {        
        Write-Log -Message "Failed to extract the HPIA: $($_.Exception.Message)" -Component "Analyze" -LogLevel 3
        $LogEntry.InstallStatus = "Failed"
        $LogEntry.Message = "Failed to extract the HPIA: $($_.Exception.Message)"
        Post-LogEntry -EntryInfo $LogEntry
        Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Failed" -Force
        Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        return
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
        Write-Log -Message "The analysis returned no recommendations. No updates are available at this time" -Component "Analyze" -LogLevel 3
        $LogEntry.InstallStatus = "Not applicable"
        $LogEntry.Message = "The analysis returned no recommendations. No updates are available at this time"
        Post-LogEntry -EntryInfo $LogEntry
        Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Not applicable" -Force
        Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        return
    }
    elseif ($Process.ExitCode -eq 4096) 
    {
        Write-Log -Message "The platform is not supported" -Component "Analyze" -LogLevel 3
        $LogEntry.InstallStatus = "Not supported"
        $LogEntry.Message = "The platform is not supported"
        Post-LogEntry -EntryInfo $LogEntry
        Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Not supported" -Force
        Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        return
    }
    Else
    {
        Write-Log -Message "HPIA exited with code $($Process.ExitCode). Expecting 0." -Component "Analyze" -LogLevel 3
        $LogEntry.InstallStatus = "HPIA error"
        $LogEntry.Message = "HPIA exited with code $($Process.ExitCode). Expecting 0."
        Post-LogEntry -EntryInfo $LogEntry
        Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "HPIA error" -Force
        Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        return
    }
}
catch 
{
    Write-Log -Message "Failed to start the HPImageAssistant.exe: $($_.Exception.Message)" -Component "Analyze" -LogLevel 3
    $LogEntry.InstallStatus = "HPIA error"
    $LogEntry.Message = "Failed to start the HPImageAssistant.exe: $($_.Exception.Message)"
    Post-LogEntry -EntryInfo $LogEntry
    Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "HPIA error" -Force
    Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
    Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    return
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
                    $Recommendation.CVAUrl = $item.Solution.Softpaq.CvaUrl
                    $Recommendation.ExeUrl = $item.Solution.Softpaq.Url
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
                    $Recommendation.CVAUrl = $item.Solution.Softpaq.CvaUrl
                    $Recommendation.ExeUrl = $item.Solution.Softpaq.Url
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
                    $Recommendation.CVAUrl = $item.Solution.Softpaq.CvaUrl
                    $Recommendation.ExeUrl = $item.Solution.Softpaq.Url
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
                    $Recommendation.CVAUrl = $item.Solution.Softpaq.CvaUrl
                    $Recommendation.ExeUrl = $item.Solution.Softpaq.Url
                    $Recommendations.Add($Recommendation)
                    Write-Log -Message ">> $($Recommendation.SoftPaqId): $($Recommendation.Name) ($($Recommendation.ReferenceVersion))" -Component "Analyze"
                }
            }

            If ($DriverRecommendations.Count -eq 0 -and $SoftwareRecommendations.Count -eq 0 -and $BIOSRecommendations.Count -eq 0 -and $FirmwareRecommendations.Count -eq 0)
            {
                Write-Log -Message "No available driver updates found at this time" -Component "Analyze" -LogLevel 3
                $LogEntry.InstallStatus = "Not applicable"
                $LogEntry.Message = "No available driver updates found at this time"
                Post-LogEntry -EntryInfo $LogEntry
                Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Not applicable" -Force
                Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
                Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
                return
            }
        }
        catch 
        {
            Write-Log -Message "Failed to parse the XML file: $($_.Exception.Message)" -Component "Analyze" -LogLevel 3
            $LogEntry.InstallStatus = "Failed"
            $LogEntry.Message = "Failed to parse the XML file: $($_.Exception.Message)"
            Post-LogEntry -EntryInfo $LogEntry
            Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Failed" -Force
            Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
            Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
    }
    Else  
    {
        Write-Log -Message "Failed to find an XML report." -Component "Analyze" -LogLevel 3
        $LogEntry.InstallStatus = "Failed"
        $LogEntry.Message = "Failed to find an XML report."
        Post-LogEntry -EntryInfo $LogEntry
        Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Failed" -Force
        Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
        Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
        return
    }
}
catch 
{
    Write-Log -Message "Failed to find an XML report." -Component "Analyze" -LogLevel 3
    $LogEntry.InstallStatus = "Failed"
    $LogEntry.Message = "Failed to find an XML report."
    Post-LogEntry -EntryInfo $LogEntry
    Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Failed" -Force
    Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
    Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue
    return
}
#endregion


###############################
## Determine what to install ##
###############################
#region
Write-Log -Message "The selected update run type is '$($UpdateRunType)'." -Component "Update"
If ($UpdateRunType -eq "All")
{
    [array]$SoftPaqNumbers = ($Recommendations | Where {$_.Type -ne "BIOS"}).SoftPaqId | Select -Unique
}
If ($UpdateRunType -eq "SpecificCategories")
{
    Write-Log -Message "The following update categories were selected: $($Categories -join ', ')." -Component "Update"
    [array]$SoftPaqNumbers = ($Recommendations | Where {$_.Type -ne "BIOS" -and $_.Type -in $Categories}).SoftPaqId | Select -Unique
}
If ($UpdateRunType -eq "CriticalOnly")
{
    [array]$SoftPaqNumbers = ($Recommendations | Where {$_.Comments -match "Critical"}).SoftPaqId | Select -Unique
}
#endregion


##########################
## Download the updates ##
##########################
#region
# This is done prior to install as some updates will kill the network connection temporarily
If ($SoftPaqNumbers.Count -ge 1)
{
    Write-Log -Message "The following SoftPaqs are marked for download: $($SoftPaqNumbers -join ', ')." -Component "Download"
    foreach ($SoftPaqNumber in $SoftPaqNumbers)
    {
        $UpdateEntry = [UpdateEntry]::new()
        $UpdateEntry.SoftPaqNumber = $SoftPaqNumber
        [array]$MatchingRecommendations = $Recommendations | Where {$_.SoftPaqId -eq $SoftPaqNumber}   
        If ($MatchingRecommendations.Count -ge 1)
        {
            # Skip BIOS updates
            If ($MatchingRecommendations[0].Type -eq "BIOS")
            {
                Write-Log -Message "$SoftPaqNumber is a BIOS update. Skipping as this solution does not perform BIOS updates." -Component "Download" -LogLevel 2
                $UpdateEntry.InstallStatus = "Skipped"
                $UpdateEntry.Message = "$SoftPaqNumber is a BIOS update. Skipping as this solution does not perform BIOS updates."
                $UpdateEntries.Add($UpdateEntry)
                Continue
            }

            Write-Log -Message "A matching recommendation was found for $SoftPaqNumber. The update is applicable" -Component "Download"
            Write-Log -Message "Downloading $SoftPaqNumber" -Component "Download"

            $UpdateEntry.ExeDownloadURL = $MatchingRecommendations[0].ExeUrl
            $UpdateEntry.ExeFilename = $UpdateEntry.ExeDownloadURL.Split('/')[-1]
            $UpdateEntry.CVADownloadURL = $MatchingRecommendations[0].CVAUrl
            $UpdateEntry.CVAFilename = $UpdateEntry.CVADownloadURL.Split('/')[-1]    
            $UpdateEntry.Category = $MatchingRecommendations[0].Type     

            # Download the CVA file
            try 
            {
                Invoke-WebRequest -Uri $UpdateEntry.CVADownloadURL -OutFile "$WorkingDirectory\$($UpdateEntry.CVAFilename)" -UseBasicParsing
                Write-Log -Message "Successfully downloaded the CVA file from $($UpdateEntry.CVADownloadURL)" -Component "Update"
            }
            catch 
            {
                Write-Log -Message "Failed to download the CVA file: $($_.Exception.Message)" -Component "Download" -LogLevel 3
                $UpdateEntry.InstallStatus = "Failed"
                $UpdateEntry.Message = "Failed to download the CVA file: $($_.Exception.Message)"
                $UpdateEntries.Add($UpdateEntry)
                Continue
            }

            # Read the CVA content
            try 
            {
                $CVAContent = Get-Content -Path "$WorkingDirectory\$($UpdateEntry.CVAFilename)" -ReadCount 0 -ErrorAction Stop
            }
            catch 
            {
                Write-Log -Message "Failed to read the CVA file: $($_.Exception.Message)" -Component "Download" -LogLevel 3
                $UpdateEntry.InstallStatus = "Failed"
                $UpdateEntry.Message = "Failed to read the CVA file: $($_.Exception.Message)"
                $UpdateEntries.Add($UpdateEntry)
                Continue
            }

            # Get the update title
            $SoftwareTitle = $CVAContent | Select-String -SimpleMatch "[Software Title]"
            If ($SoftwareTitle.Count -eq 1)
            {
                [int]$LineNumber = $SoftwareTitle.LineNumber
                $UpdateEntry.UpdateName = $CVAContent[$LineNumber].Split('=')[-1]
                Write-Log -Message "Update name: $($UpdateEntry.UpdateName)" -Component "Download"
            }

            # Get the versions
            $SoftwareVersion = $CVAContent | Select-String -SimpleMatch "Version="
            If ($SoftwareVersion.Count -ge 1)
            {
                $Version = $SoftwareVersion | Where {$_.Line.StartsWith("Version=")}
                $VendorVersion = $SoftwareVersion| Where {$_.Line.StartsWith("VendorVersion=")}
                If ($Version.Count -ge 1)
                {
                    $UpdateEntry.Version = $Version.Line.Split('=')[-1]
                    Write-Log -Message "Version: $($UpdateEntry.Version)" -Component "Download"
                }
                If ($VendorVersion.Count -ge 1)
                {
                    $UpdateEntry.VendorVersion = $VendorVersion.Line.Split('=')[-1]
                    Write-Log -Message "Vendor version: $($UpdateEntry.VendorVersion)" -Component "Download"
                }
            }

            # Get the SHA256 hash
            $SPSHAMatch = $CVAContent | Select-String -SimpleMatch "SoftPaqSHA256"
            If ($SPSHAMatch.Count -eq 1)
            {
                $UpdateEntry.SHA256Hash = $SPSHAMatch.Line.Split('=')[-1]
                Write-Log -Message "SHA256 hash: $($UpdateEntry.SHA256Hash)" -Component "Download"
            }

            # Get the silent install command
            $SilentInstallMatch = $CVAContent | Select-String -SimpleMatch "SilentInstall"
            If ($SilentInstallMatch.Count -eq 1)
            {
                #$UpdateEntry.SilentInstallCmd = $SilentInstallMatch.Line.Split('=')[-1]
                $UpdateEntry.SilentInstallCmd = $SilentInstallMatch.Line.TrimStart('SilentInstall=')
                Write-Log -Message "Silent install command: $($UpdateEntry.SilentInstallCmd)" -Component "Download"
            }

            # Build a return code list
            $ReturnCodes = [System.Collections.Generic.List[ReturnCode]]::new()
            $ReturnCodeSection = $CVAContent | Select-String -SimpleMatch "[ReturnCode]"
            If ($ReturnCodeSection.Count -eq 1)
            {
                [int]$LineNumber = $ReturnCodeSection.LineNumber
                do {
                    $NextLine = $CVAContent[$LineNumber]
                    If ($NextLine.Length -ge 1)
                    {
                        $Split = $NextLine.Split(':')
                        $ReturnCode = [ReturnCode]::new()
                        $ReturnCode.Code = $Split[0]
                        $ReturnCode.Status = $Split[1]
                        $ReturnCode.RebootRequired = $Split[2].Split('=')[0]
                        $ReturnCode.Message = $Split[2].Split('=')[-1]
                        $ReturnCodes.Add($ReturnCode)
                    }
                    $LineNumber ++
                }
                until ($NextLine.Length -eq 0)
                $UpdateEntry.ReturnCodeList = $ReturnCodes
            }

            # Download the update
            Write-Log -Message "Downloading the update from $($UpdateEntry.ExeDownloadURL)" -Component "Download"
            try 
            {
                $ExistingBitsJob = Get-BitsTransfer -Name "$($UpdateEntry.ExeFilename)" -AllUsers -ErrorAction SilentlyContinue
                If ($ExistingBitsJob)
                {
                    Write-Log -Message "An existing BITS tranfer was found. Cleaning it up." -Component "Download" -LogLevel 2
                    Remove-BitsTransfer -BitsJob $ExistingBitsJob
                }
                $BitsJob = Start-BitsTransfer -Source "http://$($UpdateEntry.ExeDownloadURL)" -Destination "$WorkingDirectory\$($UpdateEntry.ExeFilename)" -Asynchronous -DisplayName "$($UpdateEntry.ExeFilename)" -Description "$($UpdateEntry.ExeFilename) download" -RetryInterval 60 -ErrorAction Stop 
                do {
                    Start-Sleep -Seconds 5
                    $Progress = [Math]::Round((100 * ($BitsJob.BytesTransferred / $BitsJob.BytesTotal)),2)
                    Write-Log -Message "Downloaded $Progress`%" -Component "Download"
                } until ($BitsJob.JobState -in ("Transferred","Error"))
                If ($BitsJob.JobState -eq "Error")
                {
                    Write-Log -Message "BITS tranfer failed: $($BitsJob.ErrorDescription)" -Component "Download" -LogLevel 3
                    $UpdateEntry.InstallStatus = "Failed"
                    $UpdateEntry.Message = "BITS tranfer failed: $($BitsJob.ErrorDescription)"
                    $UpdateEntries.Add($UpdateEntry)
                    Continue
                }
                Write-Log -Message "Finished downloading the update" -Component "Download"
                Complete-BitsTransfer -BitsJob $BitsJob
                Write-Log -Message "BITS transfer is complete" -Component "Download"
            }
            catch 
            {
                Write-Log -Message "Failed to start a BITS transfer for the update: $($_.Exception.Message)" -Component "Download" -LogLevel 3
                $UpdateEntry.InstallStatus = "Failed"
                $UpdateEntry.Message = "Failed to start a BITS transfer for the update: $($_.Exception.Message)"
                $UpdateEntries.Add($UpdateEntry)
                Continue
            }

            # Verify the file hash
            If (Test-Path "$WorkingDirectory\$($UpdateEntry.ExeFilename)")
            {
                Write-Log -Message "Checking file hash" -Component "Download"
                try 
                {
                    $FileHash = Get-FileHash -Path "$WorkingDirectory\$($UpdateEntry.ExeFilename)" -Algorithm SHA256 -ErrorAction Stop
                    If ($FileHash.Hash -eq $UpdateEntry.SHA256Hash)
                    {
                        Write-Log -Message "The hashes match" -Component "Download"
                    }
                    else 
                    {
                        Write-Log -Message "The hash of the downloaded file does not match the expected value" -Component "Download" -LogLevel 3
                        $UpdateEntry.InstallStatus = "Failed"
                        $UpdateEntry.Message = "The hash of the downloaded file does not match the expected value"
                        $UpdateEntries.Add($UpdateEntry)
                        Continue
                    }
                }
                catch 
                {
                    Write-Log -Message "Unable to verify the hash of the downloaded file: $($_.Exception.Message)" -Component "Download" -LogLevel 3
                    $UpdateEntry.InstallStatus = "Failed"
                    $UpdateEntry.Message = "Unable to verify the hash of the downloaded file: $($_.Exception.Message)"
                    $UpdateEntries.Add($UpdateEntry)
                    Continue
                }      
            }
            else 
            {
                Write-Log -Message "Update file not found post download" -Component "Download" -LogLevel 3
                $UpdateEntry.InstallStatus = "Failed"
                $UpdateEntry.Message = "Update file not found post download"
                $UpdateEntries.Add($UpdateEntry)
                Continue
            }

            $UpdateEntries.Add($UpdateEntry)
        }
        else 
        {
            Write-Log -Message "No matching recommendation was found for SoftPaqId $SoftPaqNumber. The update is no longer applicable" -Component "Download" -LogLevel 3
            $UpdateEntry.InstallStatus = "Not applicable"
            $UpdateEntry.Message = "No matching recommendation was found for SoftPaqId $SoftPaqNumber. The update is no longer applicable"
            $UpdateEntries.Add($UpdateEntry)
            Continue
        }       
    }
}
else 
{
    Write-Log -Message "There are no applicable updates from the selection." -Component "Download" -LogLevel 2
}
#endregion


###########################
## Begin driver installs ##
###########################
#region
$UpdatesToInstall = $UpdateEntries | where {$_.InstallStatus.Length -eq 0}
If ($UpdatesToInstall.Count -ge 1)
{
    Write-Log -Message "There following SoftPaqs will be installed: $($UpdatesToInstall.SoftPaqNumber -join ', ')." -Component "Install"
    foreach ($Entry in $UpdatesToInstall)
    {
        Write-Log -Message "## $($Entry.SoftPaqNumber): $($Entry.UpdateName) ($($Entry.Version)) ##" -Component "Install"

        # Create registry key
        $FullRegPath = "$RegRoot\$ParentFolderName\$ChildFolderName\$($Entry.SoftPaqNumber)"
        $CreateRegPath = "SOFTWARE\$ParentFolderName\$ChildFolderName\$($Entry.SoftPaqNumber)"
        If (!(Test-Path $FullRegPath))
        {
            [void][Microsoft.Win32.Registry]::LocalMachine.CreateSubKey($CreateRegPath)
        }
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Running" -Force

        If ($CodeObject)
        {
            Remove-Variable -Name CodeObject -Force -ErrorAction SilentlyContinue
        }

        # If the silent install command contains more than 2 quotation marks we need to remove the first pair otherwise the cmd won't work
        $CharacterArray = $Entry.SilentInstallCmd.ToCharArray()
        $QuotationMarkCount = ($CharacterArray -eq '"').Count
        If ($QuotationMarkCount -eq 4)
        {           
            $Indexes = $CharacterArray | Select-String -SimpleMatch '"' | Select -ExpandProperty LineNumber -First 2
            $Entry.SilentInstallCmd = $Entry.SilentInstallCmd.Remove($Indexes[0]-1,1)
            $Entry.SilentInstallCmd = $Entry.SilentInstallCmd.Remove($Indexes[1]-2,1)
        }
        
        # Start the install
        Write-Log -Message "Starting the install" -Component "Install"
        try 
        {
            $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $Process = Start-Process "$WorkingDirectory\$($Entry.ExeFilename)" -ArgumentList "-s","-e cmd.exe","-a","/c $($Entry.SilentInstallCmd)" -PassThru -NoNewWindow -ErrorAction Stop
            $Handle = $Process.Handle # Needed to get the exit code when -Wait isn't used
            do {
                Start-Sleep -Seconds 1
            } 
            until ($Process.HasExited -eq $true -or $Stopwatch.Elapsed.TotalMinutes -ge $InstallTimeout)
            $Stopwatch.Stop()
            If ($Stopwatch.Elapsed.TotalMinutes -ge $InstallTimeout)
            {
                Write-Log -Message "The update installation exceeded the timeout value of $InstallTimeout minutes." -Component "Install" -LogLevel 3
                $Entry.InstallStatus = "Unknown"
                $Entry.Message = "The update installation exceeded the timeout value of $InstallTimeout minutes."
                $Entry.InstallDate = Get-Date ([DateTime]::UtcNow) -Format "s"
                Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Unknown" -Force
                Set-ItemProperty -Path $FullRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
                Continue
            }
            Write-Log -Message "The installer has finished with exit code $($Process.ExitCode)" -Component "Install"
            If ($Entry.ReturnCodeList.Code -contains $Process.ExitCode)
            {
                $CodeObject = $ReturnCodes | where {$_.Code -eq $Process.ExitCode}
                $Entry.ReturnCode = $CodeObject.Code
                $Entry.InstallStatus = $CodeObject.Status
                $Entry.RebootRequired = $CodeObject.RebootRequired
                $Entry.Message = $CodeObject.Message
                $Entry.InstallDate = Get-Date ([DateTime]::UtcNow) -Format "s"
                Write-Log -Message "Code: $($CodeObject.Code) | Status: $($CodeObject.Status) | RebootRequired: $($CodeObject.RebootRequired) | Message: $($CodeObject.Message)" -Component "Install"
                Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Complete" -Force
                Set-ItemProperty -Path $FullRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
                Continue
            }
            else 
            {
                Write-Log -Message "The return code is unexpected. The install may have failed." -Component "Install" -LogLevel 2
                $Entry.InstallStatus = "Unknown"
                $Entry.Message = "The return code is unexpected. The install may have failed."
                $Entry.ReturnCode = $Process.ExitCode
                $Entry.InstallDate = Get-Date ([DateTime]::UtcNow) -Format "s"
                Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Unknown" -Force
                Set-ItemProperty -Path $FullRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
                Continue
            }
        }
        catch 
        {
            Write-Log -Message "The installer failed: $($_.Exception.Message)" -Component "Install" -LogLevel 3
            $Entry.InstallStatus = "Failed"
            $Entry.Message = "The installer failed: $($_.Exception.Message)"
            $Entry.InstallDate = Get-Date ([DateTime]::UtcNow) -Format "s"
            Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Failed" -Force
            Set-ItemProperty -Path $FullRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
            Continue
        }             
    }
}
else 
{
    Write-Log -Message "There are no updates to install." -Component "Install" -LogLevel 2
}
#endregion


######################
## Post Log Entries ##
######################
#region
If ($UpdatesToInstall.Count -ge 1)
{
    # First test/wait for access to the workspace as some driver installs will kill network connectivity temporarily
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
        Write-Log -Message "Gave up trying to connect to the Log Analytics endpoint. The log entry will not be posted." -Component "Post" -LogLevel3
    }

    If ($tcp.Connected -eq $true)
    {
        $tcp.Close()
        $tcp.Dispose()

        foreach ($EntryInfo in $UpdatesToInstall)
        {
            Post-LogEntry -EntryInfo $EntryInfo
        }
    }
}
#endregion


###############
## Finish up ##
###############
#region
Remove-Item -Path $WorkingDirectory -Recurse -Force -ErrorAction SilentlyContinue # Clean up the working directory
Set-ItemProperty -Path $ParentRegPath -Name ExecutionStatus -Value "Completed" -Force
Set-ItemProperty -Path $ParentRegPath -Name Timestamp -Value (Get-Date ([DateTime]::UtcNow) -Format "s") -Force
Write-Log -Message "This driver update run is complete. Have a nice day!" -Component "Completion"
#endregion