############################################################
# Example Azure Automation Runbook for exporting data from #
# MS Graph and sending it to an Azure storage account      #
############################################################


# Variables
$ProgressPreference = 'SilentlyContinue'
$ResourceGroup = "<my-resource-group>" # Reource group that hosts the storage account
$StorageAccount = "<my-storage-account>" # Storage account name
$Container = "<my-container>" # Container name
$TempFolder = "$env:Temp" # Temp location to save the exported data
$CSVFileName = "Devices.csv" # Name of the exported data file



#############
# FUNCTIONS #
#############

# function to invoke a web request to MS Graph with error handling
Function script:Invoke-LocalGraphRequest {
    Param ($URL,$Headers,$Method,$Body,$ContentType)
    try {
        If ($Method -eq "Post")
        {
            $WebRequest = Invoke-WebRequest -Uri $URL -Method $Method -Headers $Headers -Body $Body -ContentType $ContentType -UseBasicParsing
        }
        else 
        {
            $WebRequest = Invoke-WebRequest -Uri $URL -Method $Method -Headers $Headers -UseBasicParsing
        }     
    }
    catch {
        $WebRequest = $_.Exception.Response
    }
    Return $WebRequest
}

# function to get managed Windows device data from MS Graph
Function Get-DeviceData {
    $URL = "https://graph.microsoft.com/beta/deviceManagement/manageddevices?`$filter=startsWith(operatingSystem,'Windows')&`$select=deviceName,Id,lastSyncDateTime,managementAgent,managementState,osVersion,skuFamily,deviceEnrollmentType,emailAddress,model,manufacturer,serialNumber,userDisplayName,joinType"
    $headers = @{'Authorization'="Bearer " + $accessToken}
    $GraphRequest = Invoke-LocalGraphRequest -URL $URL -Headers $headers -Method GET
    If ($GraphRequest.StatusCode -ne 200)
    {
        Return $GraphRequest
    }
    $JsonResponse = $GraphRequest.Content | ConvertFrom-Json
    $DeviceData = $JsonResponse.value
    If ($JsonResponse.'@odata.nextLink')
    {
        do {
            $URL = $JsonResponse.'@odata.nextLink'
            $GraphRequest = Invoke-LocalGraphRequest -URL $URL -Headers $headers -Method GET
            If ($GraphRequest.StatusCode -ne 200)
            {
                Return $GraphRequest
            }
            $JsonResponse = $GraphRequest.Content | ConvertFrom-Json
            $DeviceData += $JsonResponse.value
        } until ($null -eq $JsonResponse.'@odata.nextLink')
    }
    Return $DeviceData
}



####################
## AUTHENTICATION ##
####################

## Get MS Graph access token 
# Managed Identity
$url = $env:IDENTITY_ENDPOINT  
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]" 
$headers.Add("X-IDENTITY-HEADER", $env:IDENTITY_HEADER) 
$headers.Add("Metadata", "True") 
$body = @{resource='https://graph.microsoft.com/' } 
$script:accessToken = (Invoke-RestMethod $url -Method 'POST' -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body $body ).access_token

## Connect to Azure AD 
# Mmanaged Identity
$null = Connect-AzAccount -Identity



#########################
## THUNDERBIRDS ARE GO ##
#########################

$Devices = Get-DeviceData
$Devices | Export-Csv -Path $TempFolder\$CSVFileName -NoTypeInformation -Force
$StorageAccount = Get-AzStorageAccount -Name $StorageAccount -ResourceGroupName $ResourceGroup
try {
    $null = Set-AzStorageBlobContent -File $TempFolder\$CSVFileName -Container $Container -Blob $CSVFileName -Context $StorageAccount.Context -Force -ErrorAction Stop
}
catch {
    Write-Error -Exception $_ -Message "Failed to upload $CSVFileName to blob storage"
}