# This is really a dummy detection method that checks if the app was 'installed'
# in the last 5 minutes. Longer than this, and the app will be available for
# install again. Reason for doing this is to allow the same app to be used multiple
# times to update drivers.

# Enter the path to the HP_Driver_Updates.log below:
$InstallFile = Get-ItemProperty -Path "C:\ProgramData\Contoso\HP_Driver_Updates\HP_Driver_Updates.log" -ErrorAction SilentlyContinue
If ($InstallFile)
{
    [datetime]$LastWriteTime = $InstallFile.LastWriteTimeUtc
    $Now = [datetime]::UtcNow
    If (($Now - $LastWriteTime).TotalMinutes -lt 5)
    {
        Write-Output "Installed"
    }
}