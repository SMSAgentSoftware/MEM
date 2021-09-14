###################################################
## INVENTORY: SOFTWARE UPDATES SCHEDULED RESTART ##
###################################################

# Reboot required
If (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
{
    $RebootRequired = "True"
}
else 
{
    $RebootRequired = "False"
}

# Other locations to check for restart pending
# HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Orchestrator\RebootRequired
# HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\StateVariables | RebootRequired | 1

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

# Prepare the hash
$SoftwareUpdatesHash = @{
    SU_RebootRequired = $RebootRequired
    SU_ScheduledRebootTime = $ScheduledRebootTime
    SU_EngageReminderLastShownTime = $EngagedReminder
    SU_PendingRebootStartTime = $PendingRebootTime
}

# Convert to JSON and output
$SoftwareUpdatesJson = $SoftwareUpdatesHash | ConvertTo-Json -Compress
If ($SoftwareUpdatesJson.Length -gt 2048)
{
    Write-Output "Output is longer than the permitted length of 2048 characters."
    Exit 1
}
Else 
{
    Write-Output $SoftwareUpdatesJson
    Exit 0
}