<#PSScriptInfo
.DESCRIPTION Query the Graph Report Endpoint:- Get the total number of user mailboxes in your organization and how many are active each day of the reporting period.
https://docs.microsoft.com/en-us/graph/api/reportroot-getmailboxusagemailboxcounts?view=graph-rest-1.0

Prerequisite https://docs.microsoft.com/en-us/graph/auth-register-app-v2?context=graph%2Fapi%2F1.0&view=graph-rest-1.0
Below application permission is required:
Reports.Read.All

DCE and DCRs must be created prior.

Custom table must also be created in a Log Analytics workspace.

SPN and secret must also be created in AAD so that the runbook can be authorized to send data to the DCE.

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

#Assembly required to create bearer token to connect to DCE.
Add-Type -AssemblyName System.Web

### Step 0: Set variables required for the rest of the script.

# information needed to authenticate to AAD and obtain a bearer token so that data can be logged ot Log Analytics.
$tenantId = Get-AutomationVariable -Name'AzMonLAWTenantID' #Tenant ID the data collection endpoint resides in
$appId = Get-AutomationVariable -Name'DCEAppID' #Application ID created and granted permissions to the DCR/DCE
$appSecret = Get-AutomationVariable -Name 'DCEAppSecret' #Secret created for the appId.
$dcrUserAssignedID = "c611ebb0-61b3-4033-9eca-06e6f183fd89" #The UserAssigned Identity ObjectID assigned to the Runbook that has permission to write to the DCR.

# information needed to send data to the DCR endpoint
$dceEndpoint = Get-AutomationVariable -Name 'DCEEndpoint' #the endpoint property of the Data Collection Endpoint object
$dcrImmutableId = Get-AutomationVariable -Name 'DCRImutableID' #the immutableId property of the DCR object
$streamName = Get-AutomationVariable -Name 'StreamName' #name of the stream in the DCR that represents the destination table

### Step 1: Obtain a bearer token used later to authenticate against the DCE.
<#
try {
    Write-Output "Building token for DCE Endpoint."
    $scope = [System.Web.HttpUtility]::UrlEncode("https://monitor.azure.com//.default")   
    $body = "client_id=$appId&scope=$scope&client_secret=$appSecret&grant_type=client_credentials";
    $headers = @{"Content-Type" = "application/x-www-form-urlencoded" };
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    Write-Output "Token build SUCCESSFUL"
}
catch [System.Net.WebException] {
    Write-Warning "Exception: $($_.Exception.Message)"
    Write-Warning "FAILED creating DCE token." 
    $ErrMsg = "$_"
    Write-Output $ErrMsg
}


$bearerToken = (Invoke-RestMethod -Uri $uri -Method "Post" -Body $body -Headers $headers).access_token
#>

$scope = [System.Web.HttpUtility]::UrlEncode("https://monitor.azure.com")
$object_id = $dcrUserAssignedID # User Assigned MI object ID (principal ID)
$uri = "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=$scope&object_id=$object_id"
$response = $(Invoke-WebRequest -Uri $uri -Headers @{Metadata = "true" }) | ConvertFrom-Json
$response.access_token

### Step 2: Connect to M365 tenants to collect Outlook Activity Reports from.
$tenants = @(
    [pscustomobject]@{
        TenantName   = "Contoso";
        TenantID     = Get-AutomationVariable -Name'M365x87145483TenantID';
        ClientID     = Get-AutomationVariable -Name 'M365x87145483ClientID';
        ClientSecret = Get-AutomationVariable -Name 'M365x87145483ClientSecret';
    }<#,
    [pscustomobject]@{
        TenantName   = "CustomerName";
        TenantID     = Get-AutomationVariable -Name 'CustomerTenantID';
        ClientID     = Get-AutomationVariable -Name 'CustomerClientID';
        ClientSecret = Get-AutomationVariable -Name 'CustomerClientSecret';
    }#>
)

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
    $UTCDateTime = [DateTime]::UtcNow.ToString('u')

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
                                            
    ### Step 3: Return the results, clean up headers and convert from CSV to JSON.

    #$QueryResults
    


    #Remove special chars from header
    #$QueryResults = $QueryResultsCsv
    $QueryResults = $QueryResults.Replace('ï»¿Report Refresh Date', 'Report Refresh Date')

    #$QueryResults

    #Convert the stream result to an array  
    $ResultsArray = ConvertFrom-Csv $QueryResults
    $ResultsArray | Add-Member -NotePropertyName Tenant -NotePropertyValue $tenantName
    $ResultsArray | Add-Member -NotePropertyName TimeGenerated -NotePropertyValue $UTCDateTime
    #$ResultsArray

    #Convert to JSON
    Write-Output "Converting to JSON..."
    $JSON = ConvertTo-Json -InputObject $ResultsArray -Depth 10

    ## Can be used to create the table schema for the custom Log Analytics custom table.
    #$JSON | Set-Content M365OutlookActivityCollection.json
    #$JSON
    
    ### Step 4. Post to Log Analytics
    
    try {
        Write-Output "Posting to Log Analytics Workspace ..."
        $body = $JSON;
        $headers = @{"Authorization" = "Bearer $response.access_token"; "Content-Type" = "application/json" };
        #$headers = @{"Authorization" = "Bearer $bearerToken"; "Content-Type" = "application/json"};
        $uri = "$dceEndpoint/dataCollectionRules/$dcrImmutableId/streams/$($streamName)?api-version=2021-11-01-preview"
        $uploadResponse = Invoke-RestMethod -Uri $uri -Method "Post" -Body $body -Headers $headers
        Write-Output "Upload to Log Analytics workspace SUCCESSFUL"
    }
    catch [System.Net.WebException] {
        $SendStatus = "failure"
        Write-Warning "Exception: $($_.Exception.Message)"
        Write-Output "FAILED digesting for $tenantName" 
        $TransMsg = "$_"
        Write-Output $TransMsg
    }
}

