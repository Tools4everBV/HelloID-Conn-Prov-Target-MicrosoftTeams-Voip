#################################################
# HelloID-Conn-Prov-Target-MicrosoftTeams-Voip-Create
# PowerShell V2
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

#region functions
function Resolve-MicrosoftTeams-VoipError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $httpErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if (-not [string]::IsNullOrEmpty($ErrorObject.ErrorDetails.Message)) {
            $httpErrorObj.ErrorDetails = $ErrorObject.ErrorDetails.Message
        }
        elseif ($ErrorObject.Exception.GetType().FullName -eq 'System.Net.WebException') {
            if ($null -ne $ErrorObject.Exception.Response) {
                $streamReaderResponse = [System.IO.StreamReader]::new($ErrorObject.Exception.Response.GetResponseStream()).ReadToEnd()
                if (-not [string]::IsNullOrEmpty($streamReaderResponse)) {
                    $httpErrorObj.ErrorDetails = $streamReaderResponse
                }
            }
        }
        try {
            $errorObjectConverted = $ErrorObject | ConvertFrom-Json -ErrorAction Stop

            if ($null -ne $errorObjectConverted.error_description) {
                $httpErrorObj.FriendlyMessage = $errorObjectConverted.error_description
            }
            elseif ($null -ne $errorObjectConverted.error) {
                if ($null -ne $errorObjectConverted.error.message) {
                    $httpErrorObj.FriendlyMessage = $errorObjectConverted.error.message
                    if ($null -ne $errorObjectConverted.error.code) { 
                        $httpErrorObj.FriendlyMessage = $httpErrorObj.FriendlyMessage + " Error code: $($errorObjectConverted.error.code)"
                    }
                }
                else {
                    $httpErrorObj.FriendlyMessage = $errorObjectConverted.error
                }
            }
            else {
                $httpErrorObj.FriendlyMessage = $ErrorObject
            }
        }
        catch {
            $httpErrorObj.FriendlyMessage = $httpErrorObj.ErrorDetails
        }
        Write-Output $httpErrorObj
    }
}

function Import-ModuleIfNeeded {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [string[]]$Cmdlets
    )
    try {
        if ($Cmdlets) { 
            Import-Module $Name -Cmdlet $Cmdlets -Verbose:$false -ErrorAction Stop
        }
        else { 
            Import-Module $Name -Verbose:$false -ErrorAction Stop
        }
        Write-Information "Module [$Name] imported"
    }
    catch {
        throw "Could not load module [$Name]. Error: $_"
    }
}

function New-OAuthTokenClientSecret {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter(Mandatory)][string]$ResourceAudience
    )
    try {
        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $AppId
            client_secret = $ClientSecret
            scope         = "$ResourceAudience/.default"
        }
        $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $response = Invoke-RestMethod -Uri $uri -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
        Write-Output $response.access_token
    }
    catch {
        Throw $_
    }
}

function Connect-TeamsCertificate {
    [CmdletBinding()]
    param(
        [string[]]$EnsureCmdlets = @(
            'Get-CsOnlineUser',
            'Get-CsPhoneNumberAssignment'
        )
    )
    try {
        # Load certificate from base64 string (no permanent import needed)
        $rawCertificate = [System.Convert]::FromBase64String($actionContext.Configuration.AppCertificateBase64String)
        $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($rawCertificate, $actionContext.Configuration.AppCertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
        $thumbprint = $certificate.Thumbprint

        # Add certificate to CurrentUser\My store (required for Teams module)
        $storeRead = [System.Security.Cryptography.X509Certificates.X509Store]::new("My", "CurrentUser")
        $storeRead.Open("ReadOnly")
        $existingCert = $storeRead.Certificates | Where-Object { $_.Thumbprint -eq $thumbprint }
        $storeRead.Close()

        # Only open for write if certificate doesn't exist yet
        if ($null -eq $existingCert) {
            $storeWrite = [System.Security.Cryptography.X509Certificates.X509Store]::new("My", "CurrentUser")
            $storeWrite.Open("ReadWrite")
            $storeWrite.Add($certificate)
            $storeWrite.Close()
        }

        # Module + connect
        Import-ModuleIfNeeded -Name 'MicrosoftTeams' -Cmdlets $EnsureCmdlets
        $null = Connect-MicrosoftTeams -CertificateThumbprint $thumbprint -ApplicationId $actionContext.Configuration.AppId -TenantId $actionContext.Configuration.TenantId -ErrorAction Stop -Verbose:$false
        Write-Information 'Connected to Microsoft Teams (certificate)'
    }
    catch {
        Throw $_
    }
}
#endregion

try {
    # Initial Assignments
    $actionMessage = 'Initialization'
    $outputContext.AccountReference = 'Currently not available'

    # Validate correlation configuration
    if ($actionContext.CorrelationConfiguration.Enabled) {
        $correlationField = $actionContext.CorrelationConfiguration.AccountField
        $correlationValue = $actionContext.CorrelationConfiguration.PersonFieldValue

        if ([string]::IsNullOrEmpty($correlationField)) {
            throw 'Correlation is enabled but not configured correctly'
        }
        if ([string]::IsNullOrEmpty($correlationValue)) {
            throw 'Correlation is enabled but [accountFieldValue] is empty. Please make sure it is correctly mapped'
        }
        $actionMessage = "Connecting to Microsoft Teams"
        $null = Connect-TeamsCertificate

        $actionMessage = "Correlating account [$correlationValue] on field: [$($correlationField)] with value: [$($correlationValue)]"
        $csOnlineUser = Get-CsOnlineUser -Identity $correlationValue -ErrorAction Stop
        if (-not ($csOnlineUser.FeatureTypes -contains 'PhoneSystem' -or $csOnlineUser.ProvisionedPlan -contains 'MCOEV')) {
            throw "User [$correlationValue] has no Teams Phone (Phone System/MCOEV) license"
        }
        else {
            $properties = $outputContext.Data.PSObject.Properties.Name
            $correlatedAccount = $csOnlineUser | Select-Object -Property $properties
            Write-Information  "User [$correlationValue] has a Teams Phone (Phone System/MCOEV) license"
        }
    }
    else {
        throw 'Correlation must be enabled for this connector'
    }

    if (@($correlatedAccount).Count -eq 0) {
        $action = 'NotFound'
    }
    elseif (@($correlatedAccount).Count -eq 1) {
        $action = 'CorrelateAccount'
        $ChangeEnterpriseVoiceEnabled = ($correlatedAccount.EnterpriseVoiceEnabled -ne $actionContext.Data.EnterpriseVoiceEnabled)
    }
    elseif (@($correlatedAccount).Count -gt 1) {
        throw "Multiple accounts found for person where $correlationField is [$correlationValue]"
    }

    # Process
    switch ($action) {
        'NotFound' {
            # Throw error since creation of Teams-Voip accounts is not supported.
            $action = 'CorrelateAccount'
            throw 'Creation of MicrosoftTeams-Voip accounts is not supported. Please ensure the user receives a Teams Phone license.'
            break
        }

        'CorrelateAccount' {
            $auditMessage = "Correlated account [$($correlatedAccount.userPrincipalName)] on field [$($correlationField)] with value [$($correlationValue)]"

            if ($actionContext.DryRun -eq $true) {
                $auditMessage = "[DryRun] $auditMessage"
                Write-Information $auditMessage
            }
            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Action  = $action
                    Message = $auditMessage
                    IsError = $false
                })

            $action = 'UpdateAccount'
            if ($ChangeEnterpriseVoiceEnabled) {
                $actionMessage = "Setting Enterprise Voice to [$($actionContext.Data.EnterpriseVoiceEnabled)] for [$($correlatedAccount.userPrincipalName)]"
                $auditMessage = "Enterprise Voice set to [$($actionContext.Data.EnterpriseVoiceEnabled)] for [$($correlatedAccount.userPrincipalName)]"

                if (-not($actionContext.DryRun -eq $true)) {
                    Write-Information $actionMessage
                    $null = Set-CsPhoneNumberAssignment -Identity $correlatedAccount.Identity -EnterpriseVoiceEnabled ([bool]::Parse($actionContext.Data.EnterpriseVoiceEnabled)) -ErrorAction Stop
                }
                else {
                    $auditMessage = "[DryRun] $auditMessage"
                    Write-Information $auditMessage
                }
                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = $action
                        Message = $auditMessage
                        IsError = $false
                    })
            }

            $correlatedAccount.EnterpriseVoiceEnabled = $actionContext.Data.EnterpriseVoiceEnabled # Make sure case sensitivity is correct
            $outputContext.Data = $correlatedAccount

            $outputContext.Data.CallingLineIdentity = $correlatedAccount.CallingLineIdentity.Name
            $outputContext.AccountReference = $correlatedAccount.Identity
            $outputContext.AccountCorrelated = $true
            break
        }
    }

    $outputContext.success = $true
}
catch {
    $outputContext.success = $false
    $ex = $PSItem
    if ($($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or
        $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')) {
        $errorObj = Resolve-MicrosoftTeams-VoipError -ErrorObject $ex
        $auditMessage = "Error in action: $($actionMessage). Error: $($errorObj.FriendlyMessage)"
        Write-Warning "Error at Line '$($errorObj.ScriptLineNumber)': $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $auditMessage = "Error in action: $($actionMessage). Error: $($ex.Exception.Message)"
        Write-Warning "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    $outputContext.AuditLogs.Add([PSCustomObject]@{
            Action  = $action
            Message = $auditMessage
            IsError = $true
        })
}