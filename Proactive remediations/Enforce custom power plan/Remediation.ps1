#############################################################################
# Intune remedation script to install and / or active the named power plan  #
#############################################################################

# Variables
$PowerPlanName = "<My power plan name>"
$PowerPlanFileBase64 = "cmVnZgEAAAABAAAAxkint0hU1gAAAA=..........." # The exported power plan file in base64

# Get list of power plans
$PowerPlans = & powercfg /List

# Check that the power plan is present and active
If ($PowerPlans -match $PowerPlanName)
{
    try 
    {
        $PowerPlan = Get-CimInstance -Namespace ROOT\Cimv2\power -ClassName Win32_PowerPlan -Filter "ElementName='$PowerPlanName'" -ErrorAction Stop | 
            Sort-Object -Property IsActive -Descending | Select -First 1
        # If not active, set active
        If ($PowerPlan.IsActive -ne $true)
        {
            $GUID = $PowerPlan.InstanceID.Split('\')[1].TrimStart('{').TrimEnd('}')
            & powercfg /Setactive $GUID
        }
    }
    # Strange but true, in some cases the power plan is not present in the Win32_PowerPlan class, so fallback to powercfg
    catch 
    {
        $ActivePowerPlan = & powercfg /GetActiveScheme
        # If not active, set active
        If ($ActivePowerPlan -notmatch $PowerPlanName)
        {
            $PowerPlan = $PowerPlans | Select-String -Pattern $PowerPlanName
            $GUID = $PowerPlan.Line.Split()[3].Trim()
            & powercfg /Setactive $GUID
        }
    } 
}
# If no power plan present, import it and set as active
else 
{   
    # Save the Power plan to a file from base64  
    $PowerPlanFile = "$PowerPlanName.pow"
    $PowerPlanFileBytes = [System.Convert]::FromBase64String($PowerPlanFileBase64)
    $TempPath = [System.IO.Path]::GetTempPath()
    $PowerPlanFilePath = [System.IO.Path]::Combine($TempPath, $PowerPlanFile)
    [System.IO.File]::WriteAllBytes($PowerPlanFilePath, $PowerPlanFileBytes)
    
    # Import the power plan
    & powercfg /Import "$PowerPlanFilePath"   
    
    # Set it as active
    try 
    {
        $PowerPlan = Get-CimInstance -Namespace ROOT\Cimv2\power -ClassName Win32_PowerPlan -Filter "ElementName='$PowerPlanName'" -ErrorAction Stop | 
            Sort-Object -Property IsActive -Descending | Select -First 1
        $GUID = $PowerPlan.InstanceID.Split('\')[1].TrimStart('{').TrimEnd('}')
        & powercfg /Setactive $GUID
    }
    # Fall back to powercfg again
    catch 
    {
        $PowerPlans = & powercfg /List
        $PowerPlan = $PowerPlans | Select-String -Pattern $PowerPlanName
        $GUID = $PowerPlan.Line.Split()[3].Trim()
        & powercfg /Setactive $GUID
    } 

    # Delete the file
    Remove-Item $PowerPlanFilePath -Force
}