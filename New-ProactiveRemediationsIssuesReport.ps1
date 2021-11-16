Function New-ProactiveRemediationsIssuesReport {
<#
.Synopsis
    Creates an HTML report for devices with issues or errors in a Proactive remediations script.

.DESCRIPTION
    Runs an export job from Microsoft Graphs for a Proactive remediations script, filters out and summarises issues and errors, and creates an HTML report that can be sent by email or opened in a browser.

.PARAMETER ProactiveRemedationsScriptName
    Required. The display name of the Proactive remediations script package

.PARAMETER IncludeDevices
    Optional. Switch. Includes a list of affected devices broken down by issue or error type, in addition to a summary

.PARAMETER HTMLReport
    Optional. Switch. Generates an HTML report and opens in the default file viewer

.PARAMETER SendEmail
    Optional. Switch. Sends the HTML report by email using the mail parameter defaults defined in the function or provided as parameters

.PARAMETER To
    Optional for SendEmail. The To address.

.PARAMETER From
    Optional for SendEmail. The From address.

.PARAMETER Smtpserver
    Optional for SendEmail. The Smtpserver.

.PARAMETER Port
    Optional for SendEmail. The port.

.EXAMPLE
    Generates a PR issues report for the script package named "My PR Script" and opens it in the default file viewer for html files
    PS> .\New-ProactiveRemediationsIssuesReport -ProactiveRemedationsScriptName "My PR Script" -HTMLReport

.EXAMPLE
    Generates a PR issues report for the script package named "My PR Script" which includes data for affected devices, opens it in the default file viewer for html files and sends it by email using the email parameter defaults
    PS> .\New-ProactiveRemediationsIssuesReport -ProactiveRemedationsScriptName "My PR Script" -IncludeDevices -HTMLReport -SendEMail

.EXAMPLE
    Generates a PR issues report for the script package named "My PR Script" and sends it by email to the specified recipient using the remaining email parameter defaults
    PS> .\New-ProactiveRemediationsIssuesReport -ProactiveRemedationsScriptName "My PR Script" -SendEMail -To "joe.blow@contoso.com"

.NOTES
    Requires the Microsoft.Graph.Intune module. The user context executing the script is assumed to have the appropriate permissions in Microsoft Graph.
#>
    
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]
        $ProactiveRemedationsScriptName,
        [Parameter(Position=1)]
        [switch]
        $IncludeDevices,
        [Parameter(Position=2)]
        [switch]
        $HTMLReport,
        [Parameter(Position=3,ParameterSetName='Email')]
        [switch]
        $SendEmail,
        [Parameter(Position=4,ParameterSetName='Email',Mandatory=$false)]
        [string]
        $To = 'bill@contoso.com',
        [Parameter(Position=5,ParameterSetName='Email',Mandatory=$false)]
        [string]
        $From = 'automation@contoso.com',
        [Parameter(Position=6,ParameterSetName='Email',Mandatory=$false)]
        [string]
        $Smtpserver = 'contoso-com.mail.protection.outlook.com',
        [Parameter(Position=7,ParameterSetName='Email',Mandatory=$false)]
        [int]
        $Port = 25
    )

    # Test for the Intune PS SDK
    $IntuneSDK = Get-Module -Name Microsoft.Graph.Intune -ListAvailable -ErrorAction SilentlyContinue
    If ($null -eq $IntuneSDK)
    {
        throw "Microsoft.Graph.Intune module is required."
        Exit 1
    }

    # Get an access token
    $null = Update-MSGraphEnvironment -SchemaVersion beta -WarningAction SilentlyContinue
    $script:AccessToken = Connect-MSGraph -PassThru

    # general variables
    $ProgressPreference = 'SilentlyContinue'
    $TempFileName = "PRExportTemp"
    $Destination = "$env:TEMP"
    $HTML1 = ""
    $HTML2 = ""

    # set which report items to return
    $ReportItems = @(
        'PreRemediationDetectionScriptError'
        'PreRemediationDetectionScriptOutput'
        'PostRemediationDetectionScriptError'
        'PostRemediationDetectionScriptOutput'
        'RemediationScriptErrorDetails'
        'RemediationFailed'
        'DetectionFailed'
    )

    # Html CSS style 
$Style = @"
<style>
table { 
    border-collapse: collapse;
    font-family: sans-serif
    font-size: 10px
}
td, th { 
    border: 1px solid #ddd;
    padding: 6px;
}
th,h1 {
    padding-top: 8px;
    padding-bottom: 8px;
    text-align: left;
    background-color: #3700B3;
    color: #03DAC6
}
h3 {
    color: #ba0404;
}
</style>
"@

    # Function to make a web reqeust to Graph with error handling
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

    # Function to get an export job from Graph
    Function Get-MSGraphExportJob {
        Param($ReportName,$Filter,$FileName,$Destination)

        $bodyHash = [ordered]@{
            reportName = $ReportName
            filter = $Filter
        }
        $bodyJson = $bodyHash | ConvertTo-Json -Depth 3

        $URL = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs"
        $Headers = @{'Authorization'="Bearer " + $accessToken; 'Accept'="application/json"}
        $GraphRequest = Invoke-LocalGraphRequest -URL $URL -Headers $Headers -Method POST -Body $bodyJson -ContentType "application/json"
        If ($GraphRequest.StatusCode -ne 201)
        {
            Write-Warning "Export job Graph request returned $($GraphRequest.StatusCode)). Retrying..."
            Start-Sleep -Seconds 10
            $RetryCount = 0
            do {
                $GraphRequest = Invoke-LocalGraphRequest -URL $URL -Headers $Headers -Method POST -Body $bodyJson -ContentType "application/json"
                If ($GraphRequest.StatusCode -ne 201) 
                {
                    Write-Warning "Export job Graph request returned $($GraphRequest.StatusCode)). Retrying..."
                    $RetryCount ++
                    Start-Sleep -Seconds 10
                }
            }
            Until ($GraphRequest.StatusCode -eq 201 -or $RetryCount -ge 10)
        }
        If ($RetryCount -ge 10)
        {
            Write-Error "Gave up waiting for the export job to be created."
            throw
        }

        $Id = ($GraphRequest.Content | ConvertFrom-Json).Id
        do {
            Start-Sleep -Seconds 5
            $URL = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$Id')"
            $Headers = @{'Authorization'="Bearer " + $accessToken; 'Accept'="application/json"}
            $WebResponse = Invoke-LocalGraphRequest -URL $URL -Headers $Headers -Method GET
            $ReponseJson = $WebResponse.Content | ConvertFrom-Json
            $Status = $ReponseJson.status
        }
        Until ($Status -eq "Completed")

        $DownloadUrl = $ReponseJson.url   
        try {
            $DownloadRequest = Invoke-WebRequest -Uri $DownloadUrl -OutFile "$Destination\$FileName.zip" -UseBasicParsing -PassThru
        }
        catch {
            $DownloadRequest = $_.Exception.Response
        }
        Return $DownloadRequest

    }

    # Get the id of the PR script
    $URL = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts"
    $Headers = @{'Authorization'="Bearer " + $accessToken; 'Accept'="application/json"}
    $Response = Invoke-LocalGraphRequest -URL $URL -Headers $Headers -Method GET
    If ($Response.StatusCode -eq 200)
    {
        $Value = ($Response.Content | ConvertFrom-Json).value
        $ProactiveRemediationsScriptGUID = ($Value | where {$_.displayName -eq $ProactiveRemedationsScriptName}) | Select -ExpandProperty id -ErrorAction SilentlyContinue
        If ($null -eq $ProactiveRemediationsScriptGUID)
        {
            throw "Unable to find id for the provided script package display name"
            Exit 1
        }
    }
    else 
    {
        throw "Http error encountered from Graph API. Status code: $($Report.StatusCode). Status description: $($Report.StatusDescription)."  
        Exit 1  
    }

    # Get an export job from Graph from Proactive Remediations (DO data) and handle errors
    $Report = Get-MSGraphExportJob -ReportName DeviceRunStatesByProactiveRemediation -Filter "PolicyId eq '$ProactiveRemediationsScriptGUID'" -FileName $TempFileName -Destination $Destination
    If ($Report.GetType().BaseType -eq [System.Net.WebResponse] -or $Report.GetType().Name -eq "WebResponseObject")
    {
        If ($Report.GetType().Name -eq "WebResponseObject")
        {
            If ($Report.StatusCode -eq 504)
            {
                # Server timeout encountered, lets try again a couple of times
                Write-Warning -Message "Http 504 (gateway timeout) encountered while getting Graph data. Retrying up to 3 times."
                [int]$RetryAttempts = 0
                do {
                    $RetryAttempts ++ 
                    Start-Sleep -Seconds 5
                    $Report = Get-MSGraphExportJob -ReportName DeviceRunStatesByProactiveRemediation -Filter "PolicyId eq '$ProactiveRemediationsScriptGUID'" -FileName $FileName -Destination $Destination
                }
                until ($RetryAttempts -gt 3 -or $Report.StatusCode -eq 200)
            }
            ElseIf ($Report.StatusCode -ne 200)
            {
                throw "Http error encountered from Graph API. Status code: $($Report.StatusCode). Status description: $($Report.StatusDescription)."
                Exit 1
            }
        }
        else 
        {
            If ($Report.StatusCode.value__ -eq 504)
            {
                # Server timeout encountered, lets try again a couple of times
                Write-Warning -Message "Http 504 (gateway timeout) encountered while getting Graph data. Retrying up to 3 times."
                [int]$RetryAttempts = 0
                do {
                    $RetryAttempts ++ 
                    Start-Sleep -Seconds 5
                    $Report = Get-MSGraphExportJob -ReportName DeviceRunStatesByProactiveRemediation -Filter "PolicyId eq '$ProactiveRemediationsScriptGUID'" -FileName $FileName -Destination $Destination
                }
                until ($RetryAttempts -gt 3 -or $Report.GetType().BaseType -ne [System.Net.WebResponse])
            }
            If ($Report.GetType().BaseType -eq [System.Net.WebResponse])
            {
                throw "Http error encountered from Graph API. Status code: $($Report.StatusCode.value__). Status description: $($Report.StatusDescription)."
                Exit 1
            }
        }
    }

    # Extract the CSV from the export job and import it
    Start-Sleep -Seconds 2
    Unblock-File -Path "$Destination\$TempFileName.zip"
    $CsvFile = (Expand-Archive -Path "$Destination\$TempFileName.zip" -DestinationPath $Destination -Force -Verbose) 4>&1
    $CsvFileName = $CsvFile[-1].ToString().Split('\')[-1].Replace("'.",'')
    $File = Get-Childitem -Path $Destination\$CsvFileName -File
    Move-Item -Path $File.FullName -Destination "$Destination\$TempFileName.csv" -Force
    [array]$ImportedResults = Import-Csv $Destination\$TempFileName.csv -UseCulture


    # Process data for each report item
    $ReportItems | foreach {
        $Item = $_
        # Filter the correct data for the item type
        If ($Item -match "Error" -and $Item -match "Detection")
        {
            [array]$ScriptErrors = $ImportedResults | Where {$_.$Item.Length -ge 1 -and $_.DetectionScriptStatus -ne "Pending"}
            $ColumnHeader = "Error"
        }
        elseIf ($Item -match "Error" -and $Item -match "Remediation")
        {
            [array]$ScriptErrors = $ImportedResults | Where {$_.$Item.Length -ge 1 -and $_.DetectionScriptStatus -ne "Pending" -and $_.RemediationScriptStatus -ne "Pending"}
            $ColumnHeader = "Error"
        }
        elseif ($Item -match "Output" -and $Item -match "Detection") 
        {
            [array]$ScriptErrors = $ImportedResults | Where {$_.DetectionScriptStatus -in @("With Issues","Failed") -and $_.$Item.Length -ge 1}
            $ColumnHeader = "Output"
        }
        elseif ($Item -match "Output" -and $Item -match "Remediation") 
        {
            [array]$ScriptErrors = $ImportedResults | Where {$_.DetectionScriptStatus -ne "Pending" -and $_.RemediationScriptStatus -in @("With Issues","Failed") -and $_.$Item.Length -ge 1}
            $ColumnHeader = "Output"
        }
        elseif ($Item -eq "DetectionFailed") 
        {
            [array]$ScriptErrors = $ImportedResults | Where {$_.DetectionScriptStatus -eq "Failed"}
            $ColumnHeader = "Error"
        }
        elseif ($Item -eq "RemediationFailed") 
        {
            [array]$ScriptErrors = $ImportedResults | Where {$_.DetectionScriptStatus -ne "Pending" -and $_.RemediationScriptStatus -eq "Failed"}
            $ColumnHeader = "Error"
        }

        If ($ScriptErrors.Count -ge 1)
        {
            # DetectionFailed and RemediationFailed only
            If ($Item -in @("DetectionFailed","RemediationFailed"))
            {
                switch ($Item)
                {
                    "DetectionFailed" {$ErrorType = "Detection failed"}
                    "RemediationFailed" {$ErrorType = "Remediation failed"}
                }
                $DataTable = [System.Data.DataTable]::new()
                [void]$DataTable.Columns.AddRange(@("Count",$ColumnHeader))
                [void]$DataTable.Rows.Add($ScriptErrors.Count,$ErrorType)
                
                If ($IncludeDevices)
                {
                    If ($HTML2.Length -eq 0)
                    {
                        $HTML2 = "<H1>Devices with issues or errors</H1>"
                    }
                    $HTML2 += "<H2>$Item</H2>"
                    $HTML2 += $ScriptErrors | 
                        Select DeviceName,UserName,ModifiedTime,OSVersion,DetectionScriptStatus,RemediationScriptStatus | 
                        Sort DeviceName -Unique | 
                        ConvertTo-Html -Head $Style -PreContent "<H3>$($ErrorType)</H3>" |
                        Out-String
                }
                If ($HTML1.Length -eq 0)
                {
                    $HTML1 = "<H1>Issue summaries</H1>"
                }
                $HTML1 += $DataTable | 
                    ConvertTo-Html -Property Count,"$ColumnHeader" -Head $Style -PreContent "<H2>$Item</H2>" | 
                    Out-String           
            }
            # All others
            else 
            {
                $ScriptErrorTypes = $ScriptErrors | Group-Object -Property $Item -NoElement | Sort -Property Count -Descending
                $DataTable = [System.Data.DataTable]::new()
                [void]$DataTable.Columns.AddRange(@("Count",$ColumnHeader))

                If ($IncludeDevices)
                {
                    If ($HTML2.Length -eq 0)
                    {
                        $HTML2 = "<H1>Devices with issues or errors</H1>"
                    }
                    $HTML2 += "<H2>$Item</H2>"
                }

                foreach ($ScriptErrorType in $ScriptErrorTypes)
                {
                    [void]$DataTable.Rows.Add($ScriptErrorType.Count,$ScriptErrorType.Name)
                    
                    If ($IncludeDevices)
                    {
                        $HTML2 += $ScriptErrors | 
                            where {$_.$Item -eq $ScriptErrorType.Name} |
                            Select DeviceName,UserName,ModifiedTime,OSVersion,DetectionScriptStatus,RemediationScriptStatus | 
                            Sort DeviceName -Unique | 
                            ConvertTo-Html -Head $Style -PreContent "<H3>$($ScriptErrorType.Name)</H3>" |
                            Out-String
                    }
                }
                If ($HTML1.Length -eq 0)
                {
                    $HTML1 = "<H1>Issue summaries</H1>"
                }
                $HTML1 += $DataTable | 
                    ConvertTo-Html -Property Count,"$ColumnHeader" -Head $Style -PreContent "<H2>$Item</H2>" | 
                    Out-String
            }
        }
    }

    # Add the title
    If ($HTML1.Length -ge 1 -or $HTML2.Length -ge 1)
    {
        $HTML = "<H1>Proactive remediations issues report for '$ProactiveRemedationsScriptName'</H1>" + $HTML1 + $HTML2
    }

    If ($HTML.Length -ge 1)
    {
        # Display html report
        If ($HTMLReport)
        {
            $HTML | Out-File $Destination\$TempFileName.html -Force
            Invoke-Item $Destination\$TempFileName.html
        }
        # Email html report
        If ($SendEmail)
        {
            $EmailParams = @{
                To         = $To
                From       = $From
                Smtpserver = $Smtpserver
                Port       = $Port
                Subject    = "Proactive remedations issues report for ""$ProactiveRemedationsScriptName""  |  $(Get-Date -Format dd-MMM-yyyy)"
            }
            Send-MailMessage @EmailParams -Body $HTML -BodyAsHtml -ErrorAction Stop
        }    
    }
    Else 
    {
        Write-Output "No error data found"
    }

    # Clean up 
    Remove-Item -Path "$Destination\$TempFileName.csv" -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$Destination\$TempFileName.zip" -Force -ErrorAction SilentlyContinue
}