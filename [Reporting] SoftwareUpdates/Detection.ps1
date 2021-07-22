#################################
## INVENTORY: SOFTWARE UPDATES ##
#################################

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

# Other locations to inventory
# Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\Settings
# Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\WindowsUpdate\UX\StateVariables


$SoftwareUpdatesHash = @{
    SU_RebootRequired = $RebootRequired
    SU_ScheduledRebootTime = $ScheduledRebootTime
    SU_EngageReminderLastShownTime = $EngagedReminder
    SU_PendingRebootStartTime = $PendingRebootTime
    SU_UTCInventoryDate = [datetime]::UtcNow | Get-Date -format "yyyy-MM-ddTHH:mm:ssZ"
}

$SoftwareUpdatesJson = $SoftwareUpdatesHash | ConvertTo-Json -Compress
Write-Output $SoftwareUpdatesJson
Exit 0