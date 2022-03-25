
$NativeOSFWUpdateService = Get-CimInstance -Namespace ROOT\HP\InstrumentedBIOS -ClassName HP_BIOSEnumeration -Filter "Name='Native OS Firmware Update Service'" -ErrorAction SilentlyContinue | Select -ExpandProperty CurrentValue
If ($null -eq $NativeOSFWUpdateService)
{
    Write-output "Setting not found"
    Exit 0
}
If ($NativeOSFWUpdateService -eq "Disable")
{
    Write-Output "Setting is disabled"
    Exit 0
}
If ($NativeOSFWUpdateService -eq "Enable")
{
    # Remediation required
    Write-Output "Setting is enabled"
    Exit 1
}