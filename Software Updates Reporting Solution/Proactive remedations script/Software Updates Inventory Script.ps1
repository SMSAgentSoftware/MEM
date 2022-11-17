###########################################
## SOFTWARE UPDATE DATA GATHERING SCRIPT ##
###########################################

$FullInventorySchedule = 1 # DAYS MINIMUM number of days between sending a full inventory
$DeltaInventorySchedule = 1 #0.416 # HOURS MINIMUM number of hours between sending a delta inventory. 0.416 = 1hr
$WorkspaceID = "<WorkspaceID>" # WorkspaceID of the Log Analytics workspace
$PrimaryKey = "<PrimaryKey>" # Primary Key of the Log Analytics workspace
$ParentDirectoryName = "Contoso" # eg Company or domain name
$ProgressPreference = 'SilentlyContinue'

# Create a 'master' custom class to contain all the inventory data
class MasterClass {
    $SU_DeviceInfo
    $SU_AvailableUpdates
    $SU_UpdateLog
    $SU_MDMUpdatePolicy
    $SU_WUPolicyState
    $SU_WUPolicySettings
    $SU_WUClientInfo
    $SU_CompatMarkers
}
$MasterClass = [MasterClass]::new()

#region Functions
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
# Function to call a WU search online for available updates
Function Get-AvailableUpdatesOnline {
    $Session = New-Object -ComObject Microsoft.Update.Session           
    $Searcher = $Session.CreateUpdateSearcher()                      
    $Criteria = "IsInstalled=0"
    $SearchResult = $Searcher.Search($Criteria) 

    if ($SearchResult.Updates.Count -ge 1)
    {
        $Categories = $SearchResult.RootCategories | Select Name,CategoryID        
        [array]$Updates = $SearchResult.Updates | ForEach-Object -Process {
            if ($_.KBArticleIDs -ne $null)
            {
                $KB = 'KB' + "$($_.KBArticleIDs)"
            }
            Else 
            {
                $KB = $null
            }
            if ($_.Categories.Count -ge 1)
            {
                $CategoryID = $_.Categories[0].CategoryID 
                $Category = ($Categories | Where {$CategoryID -eq $_.CategoryID}).Name
            }
            Else
            {
                $CategoryID = $null
                $Category = $null
            }
            # Extract Windows Display version for CUs
            If ($_.Title -match "\(KB")
            {
                try 
                {
                    $SplitArray = $UpdateName.split()
                    $Index = $SplitArray.IndexOf($($SplitArray.Where({$_ -match "version"}))) + 1
                    if ($_.Title -match "Windows 11" -and $Index -eq 0)
                    {
                        $WindowsDisplayVersion = "21H2"
                    }
                    Elseif ($Index -eq 0)
                    {
                        $WindowsDisplayVersion = $null
                    }
                    else 
                    {
                        $WindowsDisplayVersion = "$($SplitArray[$Index].Trim())"
                    }
                }
                catch {}          
            }
            else 
            {
                $WindowsDisplayVersion = $null 
            }
            # Extract Windows version
            If ($_.Title -match "Windows 10")
            {
                $WindowsVersion = "Windows 10"
            }
            ElseIf ($_.Title -match "Windows 11")
            {
                $WindowsVersion = "Windows 11"
            }
            else 
            {
            $WindowsVersion = $null 
            }
            $Size = [math]::Round(($_.MaxDownloadSize / 1MB),2)
            $UpdateID = $_.Identity.UpdateID
            $HandlerID = $_.HandlerID
            $AutoSelected = $_.AutoSelectOnWebSites.ToString()
            $IsDownloaded = $_.IsDownloaded.ToString()
            $IsMandatory = $_.IsMandatory.ToString()
            $IsHidden = $_.IsHidden.ToString()
            $RebootRequired = $_.RebootRequired.ToString()
            $IsPresent = $_.IsPresent.ToString()
            $AutoSelection = $_.AutoSelection 
            $AutoDownload = $_.AutoDownload
            $BrowseOnly = $_.BrowseOnly.ToString()
            $Description = $_.Description
            If ($null -ne $_.LastDeploymentChangeTime)
            {
                $Published = (Get-Date -Date $_.LastDeploymentChangeTime)
            }
            Else
            {
                $Published = $null
            }
            New-Object -TypeName PSObject -Property @{
                Published             = $Published
                KB                    = $KB
                Name                  = $_.Title
                Category              = $Category
                Size_MB               = $Size
                UpdateID              = $UpdateID
                HandlerID             = $HandlerID
                AutoSelected          = $AutoSelected
                IsDownloaded          = $IsDownloaded 
                IsMandatory           = $IsMandatory 
                IsHidden              = $IsHidden
                RebootRequired        = $RebootRequired
                IsPresent             = $IsPresent
                AutoSelection         = $AutoSelection 
                AutoDownload          = $AutoDownload
                BrowseOnly            = $BrowseOnly
                Description           = $Description
                WindowsVersion        = $WindowsVersion
                WindowsDisplayVersion = $WindowsDisplayVersion
            }
        }
        # categorise the update
        foreach ($Update in $Updates)
        {
            If ($Update.Name -match "Feature update to")
            {
                $Update | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Feature update"
            }
            ElseIf ($Update.Name -match "Upgrade to")
            {
                $Update | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Windows upgrade"
            }
            ElseIf (($Update.Name -match "Cumulative Update for" -or $Update.Name -match "Kumulatives Update für") -and $UpdateEvent.UpdateName -notmatch ".NET Framework")
            {
                $Update | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Windows cumulative update"
            }
            ElseIf ($Update.Name -match "\(KB")
            {
                $Update | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Other Microsoft update"
            }
            ElseIf ($Update.Name -like "* - * - *")
            {
                $Update | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Driver update"
            }
            Else
            {
                $Update | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Other update"
            }
        }
        Return $Updates
    }
}
#endregion

#region Preparation
# Start a timer
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Create Reg Keys
$RegRoot = "HKLM:\SOFTWARE"
$RegParent = $ParentDirectoryName
$RegChild = "SoftwareUpdateReporting"
$FullRegPath = "$RegRoot\$RegParent\$RegChild"
If (!(Test-Path $RegRoot\$RegParent))
{
    $null = New-Item -Path $RegRoot -Name $RegParent -Force
}
If (!(Test-Path $FullRegPath))
{
    $null = New-Item -Path $RegRoot\$RegParent -Name $RegChild -Force
}

# Create local directory
$FullDirectoryPath = "$env:ProgramData\$ParentDirectoryName\$RegChild"
If (!(Test-Path -Path $FullDirectoryPath))
{
    [void][System.IO.Directory]::CreateDirectory($FullDirectoryPath)
}

# Check if an inventory is already running to avoid duplicates
$ExecutionStatus = (Get-ItemProperty -Path $FullRegPath -Name ExecutionStatus -ErrorAction SilentlyContinue).ExecutionStatus | Get-Date -ErrorAction SilentlyContinue
If ($ExecutionStatus -eq "Running")
{
    Write-Output "Another execution is currently running"
    Exit 0
}
else 
{
    Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "Running" -Force
}

# Set start time
Set-ItemProperty -Path $FullRegPath -Name LatestRunStartTime -Value (Get-Date -Format "s") -Force

# Determine inventory type for this run
$MostRecentInventoryDate = (Get-ItemProperty -Path $FullRegPath -Name MostRecentInventoryDate -ErrorAction SilentlyContinue).MostRecentInventoryDate | Get-Date -ErrorAction SilentlyContinue
$LatestFullInventoryDate = (Get-ItemProperty -Path $FullRegPath -Name LatestFullInventoryDate -ErrorAction SilentlyContinue).LatestFullInventoryDate | Get-Date -ErrorAction SilentlyContinue
$LatestDeltaInventoryDate = (Get-ItemProperty -Path $FullRegPath -Name LatestDeltaInventoryDate -ErrorAction SilentlyContinue).LatestDeltaInventoryDate | Get-Date -ErrorAction SilentlyContinue
If ($null -eq $MostRecentInventoryDate)
{
    $InventoryType = "Full"
}
If ($LatestDeltaInventoryDate -and $LatestFullInventoryDate)
{
    If (((Get-Date) - $LatestDeltaInventoryDate).TotalHours -ge $DeltaInventorySchedule -and ((Get-Date) - $LatestFullInventoryDate).TotalDays -lt $FullInventorySchedule)
    {
        $InventoryType = "Delta"
    }
    ElseIf (((Get-Date) - $LatestFullInventoryDate).TotalDays -ge $FullInventorySchedule)
    {
        $InventoryType = "Full"
    }
    Else  
    {
        Write-Output "No scheduled inventory required at this run"
        Set-ItemProperty -Path $FullRegPath -Name LatestRunResult -Value "No scheduled inventory required at this run" -Force
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "" -Force
        Exit 0
    }
}
If ($LatestDeltaInventoryDate -and !$LatestFullInventoryDate)
{
    If (((Get-Date) - $LatestDeltaInventoryDate).TotalHours -ge $DeltaInventorySchedule)
    {
        $InventoryType = "Full"
    }
    Else  
    {
        Write-Output "No scheduled inventory required at this run"
        Set-ItemProperty -Path $FullRegPath -Name LatestRunResult -Value "No scheduled inventory required at this run" -Force
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "" -Force
        Exit 0
    }
}
If (!$LatestDeltaInventoryDate -and $LatestFullInventoryDate)
{
    If (((Get-Date) - $LatestFullInventoryDate).TotalDays -ge $FullInventorySchedule)
    {
        $InventoryType = "Full"
    }
    ElseIf (((Get-Date) - $LatestFullInventoryDate).TotalHours -ge $DeltaInventorySchedule)
    {
        $InventoryType = "Delta"
    }
    Else
    {
        Write-Output "No scheduled inventory required at this run"
        Set-ItemProperty -Path $FullRegPath -Name LatestRunResult -Value "No scheduled inventory required at this run" -Force
        Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "" -Force
        Exit 0
    }
}
#endregion

#region GatherInventory
######################
## INTUNE DEVICE ID ##
######################
#region IntuneID
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
    Write-output "No Intune device Id could be found for this device"
    Set-ItemProperty -Path $FullRegPath -Name LatestRunResult -Value "No Intune device Id could be found for this device" -Force
    Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "" -Force
    Exit 1    
}
#endregion

###################
## AAD DEVICE ID ##
###################
#region AADID
$AADCert = (Get-ChildItem Cert:\Localmachine\MY | Where {$_.Issuer -match "CN=MS-Organization-Access"})
If ($null -ne $AADCert)
{
    $AADDeviceID = $AADCert.Subject.Replace('CN=','')
}
#endregion

##################
## CURRENT USER ##
##################
#region CurrentUser
# ref https://www.reddit.com/r/PowerShell/comments/7coamf/query_no_user_exists_for/
$header=@('SESSIONNAME', 'USERNAME', 'ID', 'STATE', 'TYPE', 'DEVICE')
$Sessions = query session
[array]$ActiveSessions = $Sessions | Select -Skip 1 | Where {$_ -match "Active"}
If ($ActiveSessions.Count -ge 1)
{
    $LoggedOnUsers = @()
    $indexes = $header | ForEach-Object {($Sessions[0]).IndexOf(" $_")}        
    for($row=0; $row -lt $ActiveSessions.Count; $row++)
    {
        $obj=New-Object psobject
        for($i=0; $i -lt $header.Count; $i++)
        {
            $begin=$indexes[$i]
            $end=if($i -lt $header.Count-1) {$indexes[$i+1]} else {$ActiveSessions[$row].length}
            $obj | Add-Member NoteProperty $header[$i] ($ActiveSessions[$row].substring($begin, $end-$begin)).trim()
        }
        $LoggedOnUsers += $obj
    }

    foreach ($LoggedOnUser in $LoggedOnUsers)
    {
        $LoggedOnDisplayName = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\SessionData\$($LoggedOnUser.ID)" -Name LoggedOnDisplayName -ErrorAction SilentlyContinue |
            Select -ExpandProperty LoggedOnDisplayname
        If ($LoggedOnDisplayName)
        {
            Add-Member -InputObject $LoggedOnUser -Name LoggedOnDisplayName -MemberType NoteProperty -Value $LoggedOnDisplayName -Force
        }
    }

    $LoggedOnUsersString = ($LoggedOnUsers.LoggedOnDisplayName -Join "  ||  ").Replace('[','').Replace(']','')
    $CurrentUser = $LoggedOnUsersString
}
#endregion

#################
## DEVICE INFO ##
#################
#region DeviceInfo
# OS
$Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
$CurrentMajorVersionNumber = (Get-ItemProperty -Path $Path -Name CurrentMajorVersionNumber -ErrorAction SilentlyContinue).CurrentMajorVersionNumber
$CurrentMinorVersionNumber = (Get-ItemProperty -Path $Path -Name CurrentMinorVersionNumber -ErrorAction SilentlyContinue).CurrentMinorVersionNumber
$CurrentBuild = (Get-ItemProperty -Path $Path -Name CurrentBuild -ErrorAction SilentlyContinue).CurrentBuild
$UBR = (Get-ItemProperty -Path $Path -Name UBR -ErrorAction SilentlyContinue).UBR
$DisplayVersion = (Get-ItemProperty -Path $Path -Name DisplayVersion -ErrorAction SilentlyContinue).DisplayVersion
If ($null -eq $DisplayVersion)
{
    switch ($CurrentBuild)
    {
        "19041" {$DisplayVersion = "2004"}
        "18363" {$DisplayVersion = "1909"}
        "18362" {$DisplayVersion = "1903"}
        "17763" {$DisplayVersion = "1809"}
        "17134" {$DisplayVersion = "1803"}
        "16299" {$DisplayVersion = "1709"}
        "15063" {$DisplayVersion = "1703"}
        "14393" {$DisplayVersion = "1607"}
        "10586" {$DisplayVersion = "1511"}
        "10240" {$DisplayVersion = "1507"}
        Default {$DisplayVersion = $null}

    }
}
$EditionID = (Get-ItemProperty -Path $Path -Name EditionID -ErrorAction SilentlyContinue).EditionID
$ProductName = (Get-CimInstance -ClassName Win32_OperatingSystem -Property Caption).Caption.Replace("Microsoft ",'')
# CS
$CSInfo = Get-CimInstance -ClassName Win32_ComputerSystem -Property Manufacturer,Model
# Intune
try 
{
    # This one seems the most accurate
    $LastSyncTime = Get-WinEvent -FilterHashtable @{
        LogName='Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Operational'
        ProviderName='Microsoft-Windows-DeviceManagement-Pushrouter'
        Id = 300
    } -ErrorAction Stop | 
        Where {$_.Message -match "The operation completed successfully"} | 
        Select -First 1 -ExpandProperty TimeCreated |
        Get-Date -Format "s" -ErrorAction SilentlyContinue
}
catch 
{
    $_.Exception.Message
    try 
    {
        # Fallback to this, not as accurate though
        $LastSyncTime = Get-WinEvent -FilterHashtable @{
            LogName='Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin'
            ProviderName='Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider'
            Id = 209
        } -ErrorAction Stop | 
            Where {$_.Message -match "The operation completed successfully"} | 
            Select -First 1 -ExpandProperty TimeCreated |
            Get-Date -Format "s" -ErrorAction SilentlyContinue
    }
    catch 
    {
        $LastSyncTime = $null
        $_.Exception.Message
    }
}

<# possible alternative for LastSyncTime
$LastSyncTime = Get-WinEvent -FilterHashtable @{
    LogName='Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin'
    ProviderName='Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider'
    Id = 209
} -ErrorAction Stop | 
    Where {$_.Message -match "The operation completed successfully"} | 
    Select -First 1 -ExpandProperty TimeCreated
#>

$DeviceInfo = @{}
$DeviceInfo.FullBuildNmber = "$CurrentMajorVersionNumber.$CurrentMinorVersionNumber.$CurrentBuild.$UBR"
$DeviceInfo.CurrentBuildNumber = $CurrentBuild
$DeviceInfo.CurrentPatchLevel = "$CurrentBuild.$UBR"
$DeviceInfo.DisplayVersion = $DisplayVersion
$DeviceInfo.EditionID = $EditionID
$DeviceInfo.ProductName = $ProductName
$DeviceInfo.FriendlyOSName = "$ProductName $DisplayVersion"
$DeviceInfo.Manufacturer = $CSInfo.Manufacturer
$DeviceInfo.Model = $CSInfo.Model
$DeviceInfo.LastSyncTime = $LastSyncTime
$DeviceInfo.CurrentUser = $CurrentUser

$MasterClass.SU_DeviceInfo = $DeviceInfo
#endregion


#######################
## AVAILABLE UPDATES ##
#######################
#region Availableupdates
$AvailableUpdates = Get-AvailableUpdatesOnline
$MasterClass.SU_AvailableUpdates = $AvailableUpdates
#endregion

################
## UPDATE LOG ##
################
#region UpdateLog
# Get WindowsUpdateClient events
[array]$UpdateEvents = Get-WinEvent -FilterHashtable @{
    LogName='System'
    ProviderName='Microsoft-Windows-WindowsUpdateClient'
} -ErrorAction Continue

If ($UpdateEvents.Count -ge 1)
{   
    [array]$UpdateEventArray = @()
    # process each event
    foreach ($UpdateEvent in $UpdateEvents)
    {
        # convert to xml
        [xml]$EventXML = $UpdateEvent.ToXml()

        # convert eventdata to hashtable
        $EventData = @{}
        foreach ($item in $EventXML.Event.EventData.Data)
        {
            $EventData."$($Item.Name)" = $Item.'#text'
        }

        # Extract common entries
        If ($EventData.updateList)
        {
            $UpdateName = $EventData.updateList
        }
        If ($EventData.updateTitle)
        {
            $UpdateName = $EventData.updateTitle
        }
        If ($EventData.updateGuid)
        {
            $UpdateGuid = $EventData.updateGuid.Replace('{','').Replace('}','')
        }
        If ($EventData.errorCode)
        {
            $ErrorCode = $EventData.errorCode
        }
        If ($ErrorCode)
        {
            $ErrorDescription = try{([ComponentModel.Win32Exception][int]$ErrorCode).Message}catch{$null}
        }
        If ($EventData.serviceGuid)
        {
            $ServiceGuid = $EventData.serviceGuid.Replace('{','').Replace('}','')
        }

        # Extract KB number
        If ($UpdateName -match "\(KB")
        {
            $KB = ($UpdateName.Split() | Where {$_ -match "\(KB"}).Replace("(",'').Replace(")",'').Trim()
        }
        else 
        {
            $KB = $null    
        }

        # Extract Windows Display version for CUs
        If ($UpdateName -match "\(KB")
        {
            $SplitArray = $UpdateName.split()
            $Index = $SplitArray.IndexOf($($SplitArray.Where({$_ -match "version"}))) + 1
            if ($UpdateName -match "Windows 11" -and $Index -eq 0)
            {
                $WindowsDisplayVersion = "21H2"
            }
            Elseif ($Index -eq 0)
            {
                $WindowsDisplayVersion = $null
            }
            else 
            {
                $WindowsDisplayVersion = "$($SplitArray[$Index].Trim())"
            }
        }
        else 
        {
            $WindowsDisplayVersion = $null 
        }

        # Extract Windows version
        If ($UpdateName -match "Windows 10")
        {
            $WindowsVersion = "Windows 10"
        }
        ElseIf ($UpdateName -match "Windows 11")
        {
            $WindowsVersion = "Windows 11"
        }
        else 
        {
        $WindowsVersion = $null 
        }

        # Add the event to a new array
        [array]$UpdateEventArray += [PSCustomObject]@{
            TimeCreated = Get-Date $UpdateEvent.TimeCreated.ToString() -Format "s"
            KeyWord1 = $UpdateEvent.KeywordsDisplayNames[0]
            KeyWord2 = $UpdateEvent.KeywordsDisplayNames[1]
            EventId = $UpdateEvent.Id
            ServiceGuid = $ServiceGuid
            UpdateName = $UpdateName
            KB = $KB
            UpdateId = $UpdateGuid
            ErrorCode = $ErrorCode
            ErrorDescription = $ErrorDescription
            RebootRequired = $null
            WindowsVersion = $WindowsVersion
            WindowsDisplayVersion = $WindowsDisplayVersion
        }
        Remove-Variable ErrorCode -ErrorAction SilentlyContinue
        Remove-Variable ErrorDescription -ErrorAction SilentlyContinue
    }

    # Get ServiceID list
    $ServiceIDs = (New-Object -ComObject Microsoft.Update.ServiceManager).Services | Select Name,ServiceId

    # Add some calculated properties to each event
    foreach ($UpdateEvent in $UpdateEventArray)
    {
        $UpdateEvent | Add-Member -MemberType NoteProperty -Name "ServiceName" -Value ($ServiceIDs | where {$_.ServiceID -eq $UpdateEvent.ServiceGuid}).Name
        # categorise the updates
        If ($UpdateEvent.UpdateName -match "Feature update to")
        {
            $UpdateEvent | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Feature update"
        }
        ElseIf ($UpdateEvent.UpdateName -match "Upgrade to")
        {
            $UpdateEvent | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Windows upgrade"
        }
        ElseIf ($UpdateEvent.UpdateName -match "Cumulative Update for" -and $UpdateEvent.UpdateName -notmatch ".NET Framework")
        {
            $UpdateEvent | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Windows cumulative update"
        }
        ElseIf ($UpdateEvent.UpdateName -match "\(KB")
        {
            $UpdateEvent | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Other Microsoft update"
        }
        ElseIf ($UpdateEvent.UpdateName -like "* - * - *")
        {
            $UpdateEvent | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Driver update"
        }
        Else
        {
            $UpdateEvent | Add-Member -MemberType NoteProperty -Name "UpdateType" -Value "Other update"
        }
    }
}

# remove Windows Store updates - they are numerous
[System.Collections.ArrayList]$UpdateEventArrayList = [array]($UpdateEventArray | where {!$_.UpdateName.StartsWith("9")})

# create a list of unique updates containing only the most recent entry per update
[array]$FinalEventArray = @()
$UniqueUpdateNames = $UpdateEventArrayList.UpdateName | Select -Unique
foreach ($UniqueUpdateName in $UniqueUpdateNames)
{
    $FinalEventArray += $UpdateEventArrayList.Where({$_.UpdateName -eq $UniqueUpdateName}) | Sort-Object -Property TimeCreated -Descending | Select -First 1
}

# check whether any updates are pending a reboot
$RegPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
If (Test-Path $RegPath)
{
    [array]$UpdateIdsForReboot = (Get-Item $RegPath).Property
}
If ($UpdateIdsForReboot.Count -ge 1)
{
    foreach ($UpdateIdForReboot in $UpdateIdsForReboot)
    {
        ($FinalEventArray | where {$_.UpdateId -eq $UpdateIdForReboot}).RebootRequired = "True"
    }
}

$MasterClass.SU_UpdateLog = $FinalEventArray
#endregion

#########################
## MDM UPDATE POLICIES ##
#########################
#region MDMPolicies
$Key = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update"
$Item = Get-Item -Path $Key -ErrorAction SilentlyContinue
If ($Item)
{
    $ValueNames = $Item.GetValueNames() | Where {$_ -notmatch "_ProviderSet" -and $_ -notmatch "_WinningProvider"} | Sort
    [PSCustomObject]$MDMUpdatePolicy = [Ordered]@{}
    foreach ($ValueName in $ValueNames)
    {
        $MDMUpdatePolicy.$ValueName = $Item.GetValue("$ValueName")
    }
}
$MasterClass.SU_MDMUpdatePolicy = $MDMUpdatePolicy
# Can also get policy info from WMI
# Get-CimInstance -Namespace ROOT\cimv2\mdm\dmmap -ClassName MDM_Policy_Result01_Update02
#endregion

#############################
## WU CURRENT POLICY STATE ##
#############################
#region WUPolicyState
$Key = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState"
$Item = Get-Item -Path $Key -ErrorAction SilentlyContinue
If ($Item)
{
    $ValueNames = $Item.GetValueNames() | Sort
    [PSCustomObject]$WUPolicyState = [Ordered]@{}
    foreach ($ValueName in $ValueNames)
    {
        $WUPolicyState.$ValueName = $Item.GetValue("$ValueName")
    }
}
$MasterClass.SU_WUPolicyState = $WUPolicyState
#endregion

#######################
## WU PAUSED UPDATES ##
#######################
#region WUPolicySettings
$Key1 = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\Settings"
$Key2 = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
$PausedFeatureStatus = Get-ItemProperty -Path $Key1 -Name PausedFeatureStatus -ErrorAction SilentlyContinue | Select -ExpandProperty PausedFeatureStatus
$PausedQualityStatus = Get-ItemProperty -Path $Key1 -Name PausedQualityStatus -ErrorAction SilentlyContinue | Select -ExpandProperty PausedQualityStatus
$PauseUpdatesStartTime = Get-ItemProperty -Path $Key2 -Name PauseUpdatesStartTime -ErrorAction SilentlyContinue | Select -ExpandProperty PauseUpdatesStartTime
$PauseUpdatesExpiryTime = Get-ItemProperty -Path $Key2 -Name PauseUpdatesExpiryTime -ErrorAction SilentlyContinue | Select -ExpandProperty PauseUpdatesExpiryTime
$PauseQualityUpdatesStartTime = Get-ItemProperty -Path $Key2 -Name PauseQualityUpdatesStartTime -ErrorAction SilentlyContinue | Select -ExpandProperty PauseQualityUpdatesStartTime
$PauseQualityUpdatesEndTime = Get-ItemProperty -Path $Key2 -Name PauseQualityUpdatesEndTime -ErrorAction SilentlyContinue | Select -ExpandProperty PauseQualityUpdatesEndTime
$PauseFeatureUpdatesStartTime = Get-ItemProperty -Path $Key2 -Name PauseFeatureUpdatesStartTime -ErrorAction SilentlyContinue | Select -ExpandProperty PauseFeatureUpdatesStartTime
$PauseFeatureUpdatesEndTime = Get-ItemProperty -Path $Key2 -Name PauseFeatureUpdatesEndTime -ErrorAction SilentlyContinue | Select -ExpandProperty PauseFeatureUpdatesEndTime
[PSCustomObject]$WUPolicySettings = [Ordered]@{
    PausedFeatureStatus = $PausedFeatureStatus
    PausedQualityStatus = $PausedQualityStatus
    PauseUpdatesStartTime = $PauseUpdatesStartTime
    PauseUpdatesExpiryTime = $PauseUpdatesExpiryTime
    PauseQualityUpdatesStartTime = $PauseQualityUpdatesStartTime
    PauseQualityUpdatesEndTime = $PauseQualityUpdatesEndTime
    PauseFeatureUpdatesStartTime = $PauseFeatureUpdatesStartTime
    PauseFeatureUpdatesEndTime = $PauseFeatureUpdatesEndTime
}
$MasterClass.SU_WUPolicySettings = $WUPolicySettings
#endregion

#########################
## WINDOWS UPDATE INFO ##
#########################
#region WUInfo
# Reboot required
If (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
{
    $RebootRequired = "True"
}
else 
{
    $RebootRequired = "False"
}

# ScheduledRebootTime
$RegScheduledReboot = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\StateVariables -Name ScheduledRebootTime -ErrorAction SilentlyContinue | Select -ExpandProperty ScheduledRebootTime
If ($RegScheduledReboot)
{
    $ScheduledRebootTime = [DateTime]::FromFileTimeUtc($RegScheduledReboot) | Get-Date -format "yyyy-MM-ddTHH:mm:ssZ"
}
else 
{
    $ScheduledRebootTime = $null
}

# EngageReminderLastShownTime
$RegEngagedReminder = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings -Name EngageReminderLastShownTime -ErrorAction SilentlyContinue | Select -ExpandProperty EngageReminderLastShownTime
If ($RegEngagedReminder)
{
    $EngagedReminder = $RegEngagedReminder
}
else 
{
    $EngagedReminder = $null
}

# PendingRebootStartTime
$RegPendingRebootTime = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings -Name PendingRebootStartTime -ErrorAction SilentlyContinue | Select -ExpandProperty PendingRebootStartTime
If ($RegPendingRebootTime)
{
    $PendingRebootTime = $RegPendingRebootTime
}
else 
{
    $PendingRebootTime = $null
}

# WU Service Startup Type
$WUUpdateServiceStartUpType = Get-Service -Name wuauserv -ErrorAction SilentlyContinue | Select -ExpandProperty StartType
If ($WUUpdateServiceStartUpType)
{
    $WUStartupType = $WUUpdateServiceStartUpType.ToString()
}
else 
{
    $WUStartupType = $null
}

# AutoUpdate status
$NoAutoUpdate = Get-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU -Name NoAutoUpdate -ErrorAction SilentlyContinue | Select -ExpandProperty NoAutoUpdate
If ($null -ne $NoAutoUpdate)
{
    If ($NoAutoUpdate -eq 1)
    {
        $AutoUpdate = "Disabled"
    }
    ElseIf ($NoAutoUpdate -eq 0)
    {
        $AutoUpdate = "Enabled"
    }
    else 
    {
        $AutoUpdate = $NoAutoUpdate.ToString()
    }
}
else 
{
    $AutoUpdate = "Not set"
}

# Auto Reboot with logged on users
$NoAutoRebootWithLoggedOnUsers = Get-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU -Name NoAutoRebootWithLoggedOnUsers -ErrorAction SilentlyContinue | Select -ExpandProperty NoAutoRebootWithLoggedOnUsers
If ($null -ne $NoAutoRebootWithLoggedOnUsers)
{
    If ($NoAutoRebootWithLoggedOnUsers -eq 1)
    {
        $AutoRebootLoggedOnUser = "Enabled"
    }
    ElseIf ($NoAutoRebootWithLoggedOnUsers -eq 0)
    {
        $AutoRebootLoggedOnUser = "Disabled"
    }
    else 
    {
        $AutoRebootLoggedOnUser = $NoAutoRebootWithLoggedOnUsers.ToString()
    }
}
else 
{
    $AutoRebootLoggedOnUser = "Not set"
}

$WindowsUpdateInfo = @{
    RebootRequired = $RebootRequired
    ScheduledRebootTime = $ScheduledRebootTime
    EngageReminderLastShownTime = $EngagedReminder
    PendingRebootStartTime = $PendingRebootTime
    WUServiceStartupType = $WUStartupType
    AutoUpdateStatus = $AutoUpdate
    NoAutoRebootWithLoggedOnUsers = $AutoRebootLoggedOnUser
}
$MasterClass.SU_WUClientInfo = $WindowsUpdateInfo
#endregion

########################
## COMPATIBILITY DATA ##
########################
#region CompatMarkers
$Key = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\TargetVersionUpgradeExperienceIndicators"
$Item = Get-Item -Path $Key
$SubKeyNames = $Item.GetSubKeyNames() | Where {$_ -ne "UNV"} | Sort
if ($SubKeyNames)
{
    [PSCustomObject]$CompatMarkers = [Ordered]@{}
    foreach ($SubKeyName in $SubKeyNames)
    {
        $SubKey = $Key + "\$SubKeyName"
        $RedReason = (Get-ItemProperty -Path $SubKey -Name "RedReason" -ErrorAction SilentlyContinue | Select -ExpandProperty RedReason) -join ", "
        $GatedBlockId = (Get-ItemProperty -Path $SubKey -Name "GatedBlockId" -ErrorAction SilentlyContinue | Select -ExpandProperty GatedBlockId) -join ", "
        $CompatMarkers."$($SubKeyName + '.RedReason')" = $RedReason
        $CompatMarkers."$($SubKeyName + '.GatedBlockId')" = $GatedBlockId
        If ($RedReason -ne "None" -or $GatedBlockId -ne "None")
        {
            $CompatMarkers."$($SubKeyName + '.BlockedFromUpgrade')" = "Yes"
        }
        else 
        {
            $CompatMarkers."$($SubKeyName + '.BlockedFromUpgrade')" = "No"
        }
    }
}
$MasterClass.SU_CompatMarkers = $CompatMarkers
#endregion

#endregion

#region PrepareInventory
################################################
## PREPARE THE FULL OR DELTA INVENTORY FILE/S ##
################################################
$JsonBase = ConvertTo-Json $MasterClass
# If DELTA
If ($InventoryType -eq "Delta")
{
    $LatestFullInventoryFile = [array](Get-ChildItem -Path $FullDirectoryPath\*.json -Filter *Full*) | Sort -Property LastWriteTime -Descending | Select -First 1
    If ($LatestFullInventoryFile)
    {
        # Read in the most recent full inventory data
        $ImportedJson = Get-Content $LatestFullInventoryFile.FullName -Encoding UTF8 -ReadCount 0
             
        # Test that the previous full JSON can be converted successfully
        try 
        {
            $null = $ImportedJson | ConvertFrom-Json -ErrorAction Stop
            
        }
        catch 
        {
            $SendFullInventoryPlease = $true
        }
        
        # Previous full inventory json is bad, lets send a new full one
        If ($SendFullInventoryPlease -eq $true)
        {
            $InventoryType = "Full"

            # Output Json without dates
            $Json = ConvertTo-Json $MasterClass
            $Json | Out-File -FilePath "$FullDirectoryPath\$IntuneDeviceID-Full.json" -Encoding utf8 -Force

            # Add in the dates
            Set-ItemProperty -Path $FullRegPath -Name LatestFullInventoryDate -Value (Get-Date -Format "s") -Force
            Set-ItemProperty -Path $FullRegPath -Name MostRecentInventoryDate -Value (Get-Date -Format "s") -Force
            $Stopwatch.Stop()
            $DeviceInfo.LatestInventoryType = $InventoryType
            $DeviceInfo.LatestFullInventory = (Get-ItemProperty -Path $FullRegPath -Name LatestFullInventoryDate -ErrorAction SilentlyContinue).LatestFullInventoryDate
            $DeviceInfo.LatestDeltaInventory = (Get-ItemProperty -Path $FullRegPath -Name LatestDeltaInventoryDate -ErrorAction SilentlyContinue).LatestDeltaInventoryDate
            $DeviceInfo.InventoryExecutionDuration = [math]::round($Stopwatch.Elapsed.TotalSeconds,2) 
            $MasterClass.SU_DeviceInfo = $DeviceInfo

            # Create the final Json with dates, to send
            $Json = ConvertTo-Json $MasterClass
        }
        else 
        {
            # Compare the two - if there are differences
            If ((ConvertTo-Json ($ImportedJson | ConvertFrom-Json) -Compress) -ne (ConvertTo-Json ($JsonBase | ConvertFrom-Json) -Compress)) 
            {       
                $CurrentMasterClass = [MasterClass]::new()
                $ImportedMasterClass = $ImportedJson | ConvertFrom-Json
                $Categories = $CurrentMasterClass | Get-Member -MemberType Property | Select -ExpandProperty Name | Sort
                foreach ($Category in $Categories)
                {
                    $CurrentMasterClass.$Category = ConvertTo-Json $MasterClass.$Category -Compress
                }
                $ChangedInventorySections = @()
                foreach ($Category in $Categories)
                {
                    # This adds any new categories
                    If ($null -eq $ImportedMasterClass.$Category)
                    {
                        $ChangedInventorySections += $Category
                    }
                    # This adds any changed existing categories
                    Else 
                    {
                        try {
                            $Comparison = Compare-Object (ConvertTo-Json $ImportedMasterClass.$Category -Compress -ErrorAction Stop) $CurrentMasterClass.$Category -PassThru -ErrorAction Stop | where {$_.SideIndicator -eq "=>"}
                        }
                        catch {}
                        
                        If ($null -ne $Comparison)
                        {
                            $ChangedInventorySections += $Category
                        }  
                    }          
                }
                
                [array]$MasterClassInventorySections = $MasterClass | Get-Member -MemberType Property | Select -ExpandProperty Name
                $DeltaMasterClass = [MasterClass]::new()
                foreach ($Section in $MasterClassInventorySections)
                {
                    If ($Section -in $ChangedInventorySections)
                    {
                        $DeltaMasterClass.$Section = $MasterClass.$Section
                    }
                }

                # Output the Delta without dates
                $DeltaJson = ConvertTo-Json $DeltaMasterClass
                $DeltaJson | Out-File -FilePath "$FullDirectoryPath\$IntuneDeviceID-$InventoryType.json" -Encoding utf8 -Force

                # Output the Full without dates, will become the 'baseline' inventory for next run
                $Json = ConvertTo-Json $MasterClass
                $Json | Out-File -FilePath "$FullDirectoryPath\$IntuneDeviceID-Full.json" -Encoding utf8 -Force
        
                # Add in the dates
                Set-ItemProperty -Path $FullRegPath -Name LatestDeltaInventoryDate -Value (Get-Date -Format "s") -Force
                Set-ItemProperty -Path $FullRegPath -Name MostRecentInventoryDate -Value (Get-Date -Format "s") -Force
                $Stopwatch.Stop()
                $DeviceInfo.LatestInventoryType = $InventoryType
                $DeviceInfo.LatestFullInventory = (Get-ItemProperty -Path $FullRegPath -Name LatestFullInventoryDate -ErrorAction SilentlyContinue).LatestFullInventoryDate
                $DeviceInfo.LatestDeltaInventory = (Get-ItemProperty -Path $FullRegPath -Name LatestDeltaInventoryDate -ErrorAction SilentlyContinue).LatestDeltaInventoryDate
                $DeviceInfo.InventoryExecutionDuration = [math]::round($Stopwatch.Elapsed.TotalSeconds,2) 
                $MasterClass.SU_DeviceInfo = $DeviceInfo       

                # Create the final delta Json with dates, to send
                $Json = ConvertTo-Json $DeltaMasterClass
            }
            # If no differences, gracefully exit
            else 
            {
                $Stopwatch.Stop()
                Write-Output "Delta contains no new data" 
                Set-ItemProperty -Path $FullRegPath -Name LatestRunResult -Value "Delta contains no new data" -Force
                Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "" -Force
                Exit 0
            }
        }
    }
    # No previous full inventory file found, therefore converting delta to full
    Else  
    {
        $InventoryType = "Full"

        # Output Json without dates
        $Json = ConvertTo-Json $MasterClass
        $Json | Out-File -FilePath "$FullDirectoryPath\$IntuneDeviceID-Full.json" -Encoding utf8 -Force

        # Add in the dates
        Set-ItemProperty -Path $FullRegPath -Name LatestFullInventoryDate -Value (Get-Date -Format "s") -Force
        Set-ItemProperty -Path $FullRegPath -Name MostRecentInventoryDate -Value (Get-Date -Format "s") -Force
        $Stopwatch.Stop()
        $DeviceInfo.LatestInventoryType = $InventoryType
        $DeviceInfo.LatestFullInventory = (Get-ItemProperty -Path $FullRegPath -Name LatestFullInventoryDate -ErrorAction SilentlyContinue).LatestFullInventoryDate
        $DeviceInfo.LatestDeltaInventory = (Get-ItemProperty -Path $FullRegPath -Name LatestDeltaInventoryDate -ErrorAction SilentlyContinue).LatestDeltaInventoryDate
        $DeviceInfo.InventoryExecutionDuration = [math]::round($Stopwatch.Elapsed.TotalSeconds,2) 
        $MasterClass.SU_DeviceInfo = $DeviceInfo

        # Create the final Json with dates, to send
        $Json = ConvertTo-Json $MasterClass
    }
}
# If FULL
Else  
{   
    # Output Json without dates
    $Json = ConvertTo-Json $MasterClass
    $Json | Out-File -FilePath "$FullDirectoryPath\$IntuneDeviceID-Full.json" -Encoding utf8 -Force

    # Add in the dates
    Set-ItemProperty -Path $FullRegPath -Name LatestFullInventoryDate -Value (Get-Date -Format "s") -Force
    Set-ItemProperty -Path $FullRegPath -Name MostRecentInventoryDate -Value (Get-Date -Format "s") -Force
    $Stopwatch.Stop()
    $DeviceInfo.LatestInventoryType = $InventoryType
    $DeviceInfo.LatestFullInventory = (Get-ItemProperty -Path $FullRegPath -Name LatestFullInventoryDate -ErrorAction SilentlyContinue).LatestFullInventoryDate
    $DeviceInfo.LatestDeltaInventory = (Get-ItemProperty -Path $FullRegPath -Name LatestDeltaInventoryDate -ErrorAction SilentlyContinue).LatestDeltaInventoryDate
    $DeviceInfo.InventoryExecutionDuration = [math]::round($Stopwatch.Elapsed.TotalSeconds,2) 
    $MasterClass.SU_DeviceInfo = $DeviceInfo

    # Create the final Json with dates
    $Json = ConvertTo-Json $MasterClass
}

#endregion

#region PostInventory
###################################
## SEND THE DATA TO LA WORKSPACE ##
###################################
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$MasterJson = ConvertFrom-Json $Json
$Categories = ($MasterJson | Get-Member -MemberType NoteProperty).Name
$SuccessfullySent = 0
$UnsuccessfullySent = 0
$UnsuccessfulItems = @()
$InventoryDate = Get-Date ([DateTime]::UtcNow) -Format "s"
foreach ($Category in $Categories)
{
    $CategoryData = $MasterJson.$Category
    If ($null -ne $CategoryData)
    {
        If ($CategoryData.GetType().BaseType.Name -eq "Array")
        {
            foreach ($item in $CategoryData)
            {
                $item | Add-Member -MemberType NoteProperty -Name IntuneDeviceID -Value $IntuneDeviceID -Force
                $item | Add-Member -MemberType NoteProperty -Name AADDeviceID -Value $AADDeviceID -Force
                $Item | Add-Member -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME -Force
                $Item | Add-Member -MemberType NoteProperty -Name InventoryDate -Value $InventoryDate -Force
            }
        }
        else 
        {
            $CategoryData | Add-Member -MemberType NoteProperty -Name IntuneDeviceID -Value $IntuneDeviceID -Force
            $CategoryData | Add-Member -MemberType NoteProperty -Name AADDeviceID -Value $AADDeviceID -Force
            $CategoryData | Add-Member -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME -Force
            $CategoryData | Add-Member -MemberType NoteProperty -Name InventoryDate -Value $InventoryDate -Force
        }
        $CategoryJson = $CategoryData | ConvertTo-Json -Compress
        $Result = Post-LogAnalyticsData -customerId $WorkspaceID -sharedKey $PrimaryKey -body ([System.Text.Encoding]::UTF8.GetBytes($CategoryJson)) -logType $Category
        If ($Result.GetType().Name -eq "WebResponseObject")
        {
            If ($Result.StatusCode -eq 200)
            {
                $SuccessfullySent ++
            }
            else 
            {
                $UnsuccessfullySent ++
                If ($null -ne $Result.Response)
                {
                    $UnsuccessfulItems += [PSCustomObject]@{
                        Category = $Category
                        StatusCode = $Result.Response.StatusCode
                        StatusDescription = $Result.Response.StatusDescription
                    }          
                }
                Else 
                {
                    $UnsuccessfulItems += [PSCustomObject]@{
                        Category = $Category
                        StatusCode = $null
                        StatusDescription = $Result.Message
                    }
                }
            }
        }
        ElseIf ($Result.GetType().Name -eq "ErrorRecord")
        {
            $UnsuccessfullySent ++
            $UnsuccessfulItems += [PSCustomObject]@{
                Category = $Category
                StatusCode = $null
                StatusDescription = $Result.Exception.Message
            }
        }
        Else 
        {
            $UnsuccessfullySent ++
            $UnsuccessfulItems += [PSCustomObject]@{
                Category = $Category
                StatusCode = $null
                StatusDescription = $Result
            }
        }
    }
}

If ($UnsuccessfullySent -ge 1)
{
    Set-ItemProperty -Path $FullRegPath -Name LatestRunResult -Value "$InventoryType inventory failed to send all categories. Successfully sent: $SuccessfullySent. Unsuccessfully sent: $UnsuccessfullySent)" -Force
    Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "" -Force
    Write-Output $UnsuccessfulItems
    Exit 1
}
else 
{
    Set-ItemProperty -Path $FullRegPath -Name LatestRunResult -Value "$InventoryType inventory sent. Category count: $SuccessfullySent" -Force
    Write-Output "$InventoryType inventory sent at $(Get-Date -Format ""s""). Category count: $SuccessfullySent"
    Set-ItemProperty -Path $FullRegPath -Name ExecutionStatus -Value "" -Force
}
#endregion