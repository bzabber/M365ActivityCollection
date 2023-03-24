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
#./Get-OndriveUsageRpt.ps1 -TenantID 'tenatId' -ClientID 'clientId' -ClientSecret '********' | Out-File -FilePath Logoutput.txt
param (
    $TenantId,
    $ClientId,
    $ClientSecret
)

# OAuth Token Endpoint
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
#Write-Host $uri
# Construct Body for OAuth Token
$body = @{
    client_id     = $ClientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
}

# Get Token
$TokenRequest = try {
    Invoke-RestMethod -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop
}
catch [System.Net.WebException] {	
    Write-Warning "Exception was caught: $($_.Exception.Message)"
}

If ($TokenRequest) { $token = $TokenRequest.access_token }
Else { "No token detected. Ending."; Break }
#write-host $token
   
       
# Create an empty array to store the result.
$QueryResults = @()
# Invoke REST method and fetch data until there are no pages left.
$requestURI = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageMailboxCounts(period='D7')"                
do {
    $Results = try {
        Invoke-RestMethod -Method Get -Uri $requestURI -ContentType "application/json" -Headers @{ Authorization = "Bearer $token" } -ErrorAction Stop
    }
    catch [System.Net.WebException] {
        Write-Warning "Exception: $($_.Exception.Message)"
    }
    if ($Results.value) {
        $QueryResults += $Results.value
                                                                    
    }
    else {
        $QueryResults += $Results
    }
    $requestURI = $Results.'@odata.nextlink'
                                                                                                
} until (!($requestURI))
                                            
# Return the resutls

$QueryResults
	           

#Remove special chars from header
$QueryResults = $QueryResults.Replace('ï»¿Report Refresh Date', 'Report Refresh Date')
#Convert the stream result to an array
$resultarray = ConvertFrom-Csv -InputObject $QueryResults
#Export result to CSV
$resultarray | Export-Csv "C:\Users\yogeshjain\Documents\MyDocs\GraphTeam\SupportTicket\SiteUsage.csv" -NoTypeInformation
