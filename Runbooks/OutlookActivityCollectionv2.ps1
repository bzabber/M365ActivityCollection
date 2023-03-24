<#PSScriptInfo
.DESCRIPTION Query the Graph Report Endpoint:- Get the total number of user mailboxes in your organization and how many are active each day of the reporting period.
https://docs.microsoft.com/en-us/graph/api/reportroot-getmailboxusagemailboxcounts?view=graph-rest-1.0

Prerequisite https://docs.microsoft.com/en-us/graph/auth-register-app-v2?context=graph%2Fapi%2F1.0&view=graph-rest-1.0
Below application permission is required:
Reports.Read.All

#NOTE - Disclaimer
#Following programming examples is for illustration only, without warranty either expressed or implied,
#including, but not limited to, the implied warranties of merchantability and/or fitness for a particular purpose. 
#This sample code assumes that you are familiar with the programming language being demonstrated and the tools 
#used to create and debug procedures. This sample code is provided for the purpose of illustration only and is 
#not intended to be used in a production environment. 

#>

<#
.SYNOPSIS Query Office 365 status

.PARAMETER TenantID
The Tenant ID of your Office 365 instance.

.PARAMETER ClientID
Use the App Registration ID.

.PARAMETER ClientSecret
Use the client secret you create as part of the app registration.

#>
#./Get-emailActivityUsageRpt.ps1 -TenantID 'tenantId' -ClientID 'clientId' -ClientSecret '********' | Out-File -FilePath Logoutput.txt
#Import-Module MSAL.PS
#--- Include module to format and send request to OMS ---#
#Import-Module OMSIngestionAPI

$tenants = @(
    [pscustomobject]@{
        TenantName   = "Contoso";
        TenantID     = "38aa4e2b-120e-426b-b877-29a4fb7be021" #'M365x87145483TenantID';
        ClientID     = "bf9c660f-45ab-4100-8c5a-10dd1dcebbc8" #Get-AutomationVariable -Name 'M365x87145483ClientID';
        ClientSecret = "JoF8Q~MIIwBduxqXn4yX4Sh96ZjwLZ-Cf-H7qdpN" #'M365x87145483ClientSecret';
    }<#,
    [pscustomobject]@{
        TenantName   = "SSC";
        TenantID     = Get-AutomationVariable -Name 'SSCTenantID';
        ClientID     = Get-AutomationVariable -Name 'SSCClientID';
        ClientSecret = Get-AutomationVariable -Name 'SSCClientSecret';
    }#>
)

#Log Analytics components
$LogAnalyticsWorkspaceID = "6c103a25-8ac4-4421-8a51-03b4c8082eb7" #Get-AzAutomationVariable -Name 'LAWWorkSpaceID'
$SharedKey = "AhLfQV/YgsSHXEafWJOMuSDTeWRZ2+++ErLzxJ65PBITRme/0Gh8Mk45t/f/hnOY+bvUdwbshzpGOh/S0EBcuA==" #Get-AzAutomationVariable -Name 'LAWPrimaryKey'
$LogType = "OutlookActivityCollection"

# Function to to create the authorization signature
Function Build-Signature($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource)
    {
        $xHeaders = "x-ms-date:" + $date
        $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource

        $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
        $keyBytes = [Convert]::FromBase64String($sharedKey)

        $sha256 = New-Object System.Security.Cryptography.HMACSHA256
        $sha256.Key = $keyBytes
        $calculatedHash = $sha256.ComputeHash($bytesToHash)
        $encodedHash = [Convert]::ToBase64String($calculatedHash)
        $authorization = 'SharedKey {0}:{1}' -f $customerId, $encodedHash
        return $authorization
    }

# Function to Post data to Azure Monitor Log Analytics Data Collection Endpoint
# Create the function to create and post the request
Function Post-LogAnalyticsData($customerId, $sharedKey, $body, $logType) {
    $method = "POST"
    $contentType = "application/json"
    $resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $contentLength = $body.Length
    $signature = Build-Signature `
        -customerId $customerId `
        -sharedKey $sharedKey `
        -date $rfc1123date `
        -contentLength $contentLength `
        -method $method `
        -contentType $contentType `
        -resource $resource
    $uri = "https://" + $customerId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"

    $headers = @{
        "Authorization"        = $signature;
        "Log-Type"             = $logType;
        "x-ms-date"            = $rfc1123date;
        "time-generated-field" = $TimeStampField;
    }

    $response = Invoke-WebRequest -Uri $uri -Method $method -ContentType $contentType -Headers $headers -Body $body -UseBasicParsing
    return $response.StatusCode

}

# Connect to each tenant and collect Outlook Activity.
foreach ($tenant in $tenants) {
    $tenantName = $tenant.TenantName
    $tenantID = $tenant.TenantID
    $clientID = $tenant.ClientID
    $ClientSecret = $tenant.ClientSecret
    
    Write-Output " Processing ... " $tenantName


    # OAuth Token Endpoint
    $uri = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"
    #Write-Host $uri
    # Construct Body for OAuth Token
    $body = @{
        client_id     = $clientID
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    # Get Token
    $TokenRequest = try {
        Invoke-RestMethod -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop
    }
    catch [System.Net.WebException] {	
        Write-Warning "Exception was caught: $($_.Exception.Message)" #+ "for "$tenantName
    }

    If ($TokenRequest) { $token = $TokenRequest.access_token }
    Else { "No token detected. Ending."; Break }
    #write-host $token
   
     
    # Create an empty array to store the result.
    $QueryResults = @()

    # Invoke REST method and fetch data until there are no pages left.
    $requestURI = "https://graph.microsoft.com/v1.0/reports/getEmailActivityCounts(period='D7')"                
    do {
        $Results = try {
            Invoke-RestMethod -Method Get -Uri $requestURI -ContentType "application/json" -Headers @{ Authorization = "Bearer $token" } -ErrorAction Stop
        }
        catch [System.Net.WebException] {
            Write-Warning "Exception: $($_.Exception.Message) for " $tenantName
        }
        if ($Results.value) {
            $QueryResults += $Results.value
                                                                    
        }
        else {
            $QueryResults += $Results
        }
        $requestURI = $Results.'@odata.nextlink'
                                                                                                
    } until (!($requestURI))
                                            
    # Return the results

    #$QueryResults


    #Remove special chars from header
    $QueryResults = $QueryResults.Replace('ï»¿Report Refresh Date', 'Report Refresh Date')
        
    #Convert the stream result to an array  
    $ResultsArray = ConvertFrom-Csv -InputObject $QueryResults
    $ResultsArray | Add-Member -NotePropertyName Tenant -NotePropertyValue $tenantName
    $ResultsArray

    #Convert to JSON
    $ResultsJSON = $ResultsArray | ConvertTo-Json

    $ResultsJSON
    
    Write-Output "Boo"
    #Export result to CSV
    #Write-Host $resultarray
    Post-LogAnalyticsData -customerId $customerId -sharedKey $sharedKey -body ([System.Text.Encoding]::UTF8.GetBytes($json)) -logType $logType
}