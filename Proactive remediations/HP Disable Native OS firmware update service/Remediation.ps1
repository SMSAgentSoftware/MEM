# Parameters
$BIOSPassword = "VGhpc0lzbnRNeVJlYWxQYXNzd29yZA=="
$SettingName = "Native OS Firmware Update Service"
$SettingValue = "Disable"


# Function to alter a BIOS setting
Function Set-HPBIOSSetting {    
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$false)]
        $Password,
        [Parameter(Mandatory=$true)]
        $Name,
        [Parameter(Mandatory=$true)]
        $Value

    )

    $CimInstance = Get-CimInstance -Namespace ROOT\HP\InstrumentedBIOS -ClassName HP_BIOSSettingInterface
    If ($Password)
    {
        $params = @{
            Name = "$Name"
            Value = "$Value"
            Password = "<utf-16/>$Password"
        }
    }
    Else
    {
        $params = @{
        Name = "$Name"
        Value = "$Value"
        Password = ""
        }
    }
    $Result = Invoke-CimMethod -InputObject $CimInstance -MethodName SetBIOSSetting -Arguments $params
    Switch ($Result.Return) {
        0 {$ResultDescription = "Success"}
        1 {$ResultDescription = "Not Supported"}
        2 {$ResultDescription = "Unknown Error"}
        3 {$ResultDescription = "Timeout"}
        4 {$ResultDescription = "Failed"}
        5 {$ResultDescription = "Invalid Parameter"}
        6 {$ResultDescription = "Access Denied"}
        32768 {$ResultDescription = "Security Policy is violated"}
        32769 {$ResultDescription = "Security Condition is not met"}
        32770  {$ResultDescription = "Security Configuration"}
        default {$ResultDescription = "Unknown"}
    }
    Return $ResultDescription
}

# Check if we can access the HP WMI namespace
Try
{
    $null = Get-CimClass -Namespace ROOT\HP\InstrumentedBIOS -ClassName HP_BIOSSettingInterface -ErrorAction Stop
}
Catch
{
    Write-Output "HP_BIOSSettingInterface class not found"
    Exit 1
}

# Check if a BIOS password has been set
$SetupPwd = (Get-CimInstance -Namespace ROOT\HP\InstrumentedBIOS -ClassName HP_BIOSPassword -Filter "Name='Setup Password'").IsSet
If ($SetupPwd -eq 1)
{
    $Password = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($BIOSPassword))
}

If ($Password)
{
    $Result = Set-HPBIOSSetting -Password $Password -Name $SettingName -Value $SettingValue -ErrorAction Stop
    Write-Output "$Result"
}
else 
{
    $Result = Set-HPBIOSSetting -Name $SettingName -Value $SettingValue -ErrorAction Stop
    Write-Output "$Result"
}
