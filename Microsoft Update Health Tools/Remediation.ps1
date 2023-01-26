# variables
$DownloadDirectory = "$env:Temp"
$DownloadFileName = "Expedite_packages.zip"
$LogDirectory = "$env:Temp"
$LogFile = "UpdateTools.log"

# Get the download URL
$ProgressPreference = 'SilentlyContinue'
$URL = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=103324"
$Request = Invoke-WebRequest -Uri $URL -UseBasicParsing
$DownloadURL = ($Request.Links | Where {$_.outerHTML -match "click here to download manually"}).href

# Download and extract the ZIP package
Invoke-WebRequest -Uri $DownloadURL -OutFile "$DownloadDirectory\$DownloadFileName" -UseBasicParsing
If (Test-Path "$DownloadDirectory\$DownloadFileName")
{
    Expand-Archive -Path "$DownloadDirectory\$DownloadFileName" -DestinationPath $DownloadDirectory -Force
}
else 
{
    Write-Host "Update tools not downloaded"    
    Exit 1
}

# Determine which cab to use
[int]$CurrentBuild = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "CurrentBuildNumber" | Select -ExpandProperty CurrentBuildNumber
if ($CurrentBuild -ge 22000)
{
    $OSVersion = "Windows11"
}
else 
{
    $OSVersion = "Windows10"    
}
[string]$Bitness = Get-CimInstance Win32_OperatingSystem -Property OSArchitecture | Select -ExpandProperty OSArchitecture
Switch ($Bitness)
{
    "64-bit" {$arch = "x64"}
    "32-bit" {$arch = "x86"}
    default { Write-Host "Unable to determine OS architecture"; Exit 1 }
}
If ($CurrentBuild -eq 22621)
{
    $CabLocation = "$DownloadDirectory\$($DownloadFileName.Split('.')[0])\$OSVersion\22H2\$arch"   
}
else 
{
    $CabLocation = "$DownloadDirectory\$($DownloadFileName.Split('.')[0])\$OSVersion\$arch"     
}
$CabName = (Get-ChildItem $CabLocation -Name *.cab).pschildname

# Expand the cab and get the MSI
expand.exe /r "$CabLocation\$CabName" /F:* "$DownloadDirectory\$($DownloadFileName.Split('.')[0])"
Start-Sleep -Seconds 5
expand.exe /r "$DownloadDirectory\$($DownloadFileName.Split('.')[0])\UpdHealthTools.cab" /F:* "$DownloadDirectory\$($DownloadFileName.Split('.')[0])"
Start-Sleep -Seconds 5
$File = Get-Childitem -Path "$DownloadDirectory\$($DownloadFileName.Split('.')[0])\*.msi" -File

# Install the MSI
$Process = Start-Process -FilePath msiexec.exe -ArgumentList "/i $($File.FullName) /qn REBOOT=ReallySuppress /L*V ""$LogDirectory\$LogFile""" -Wait -PassThru
Remove-Item "$DownloadDirectory\$($DownloadFileName.Split('.')[0])" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$DownloadDirectory\$DownloadFileName" -Force -ErrorAction SilentlyContinue
If ($Process.ExitCode -eq 0)
{
    Write-Host "Microsoft Update Health tools successfully installed"
    Exit 0
}
else 
{
    Write-Host "Microsoft Update Health tools installation failed with exit code $($Process.ExitCode)"    
    Exit 1
}
