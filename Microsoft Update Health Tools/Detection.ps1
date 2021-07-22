# Check minimum OS version requirement (1809 or later)
[int]$CurrentBuild = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "CurrentBuildNumber" | Select -ExpandProperty CurrentBuildNumber
If ($CurrentBuild -notin (17763,18363,19041,19042) -and $CurrentBuild -lt 19043)
{
    Write-Host "Minimum OS version requirement not met"
    Exit 0
}

# Check if Update tools installed
$Results = @()
$UninstallKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
foreach ($key in $UninstallKeys)
{
    If (Test-Path $key)
    {
        $Entry = (Get-ChildItem -Path $key) | Where {$_.GetValue('DisplayName') -eq "Microsoft Update Health Tools"} 
        If ($Entry)
        {
            $Results += [PSCustomObject]@{
                DisplayName = $Entry.GetValue("DisplayName")
                DisplayVersion = $Entry.GetValue("DisplayVersion")
                InstallDate = $Entry.GetValue("InstallDate")
                GUID = $Entry.pschildname
            }
        }
    }
}
If ($Results.Count -ge 1)
{
    foreach ($Result in $Results)
    {
        Write-Host ($Result | ConvertTo-Json -Compress)
    }
    Exit 0
}
Else 
{
    Write-Host "Update Tools not found"
    Exit 1
}