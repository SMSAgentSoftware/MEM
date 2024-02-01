##################################################################################
# Intune detection script to check if the named power plan is present and active #
##################################################################################

# Optional: Exit if the device is a virtual machine
<#
$Model = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Model 
If ($Model -match "Virtual Machine" -or $Model -match "VMWare")
{
    Write-Host "A virtual machine"
    Exit 0
}
#>

# Variables
$PowerPlanName = "<My power plan name>"

# Get list of power plans
$PowerPlans = & powercfg /List

# Check that the power plan is present
If ($PowerPlans -match $PowerPlanName)
{
    try 
    {
        $PowerPlan = Get-CimInstance -Namespace ROOT\Cimv2\power -ClassName Win32_PowerPlan -Filter "ElementName='$PowerPlanName'" -ErrorAction Stop | 
        Sort-Object -Property IsActive -Descending | Select -First 1
        # If not active, call remediation script
        If ($PowerPlan.IsActive -ne $true)
        {
            Write-Host "Power plan is present but not active"
            Exit 1
        }
    }
    # Strange but true, in some cases the power plan is not present in the Win32_PowerPlan class, so fallback to powercfg
    catch 
    {
        $ActivePowerPlan = & powercfg /GetActiveScheme
        If ($ActivePowerPlan -notmatch $PowerPlanName)
        {
            Write-Host "Power plan is present but not active"
            Exit 1
        }
    }
}
# If no power plan present, call remediation script
else 
{   
    Write-Host "Power plan is not present"
    Exit 1
}