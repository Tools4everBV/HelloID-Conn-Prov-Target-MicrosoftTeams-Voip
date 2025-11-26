#################################################
# HelloID-Conn-Prov-Target-MicrosoftTeams-Voip-Update
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
            'Get-CsCallingLineIdentity',
            'Grant-CsCallingLineIdentity'
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
    $action = 'UpdateAccount'

    # Verify if [aRef] has a value
    if ([string]::IsNullOrEmpty($actionContext.References.Account)) {
        throw 'The account reference could not be found'
    }

    Write-Information 'Verifying if a MicrosoftTeams-Voip account exists'

    $actionMessage = "Connecting to Microsoft Teams"
    $null = Connect-TeamsCertificate

    $actionMessage = "Retrieving account [$($actionContext.References.Account)]"
    $csOnlineUser = Get-CsOnlineUser -Identity $actionContext.References.Account -ErrorAction Stop
    if (-not ($csOnlineUser.FeatureTypes -contains 'PhoneSystem' -or $csOnlineUser.ProvisionedPlan -contains 'MCOEV')) {
        Write-Warning "User [$($csOnlineUser.userPrincipalName)] has no Teams Phone (Phone System/MCOEV) license"
    }
    else {
        $properties = $outputContext.Data.PSObject.Properties.Name
        $correlatedAccount = $csOnlineUser | Select-Object -Property $properties -ExcludeProperty DepartmentId
        Write-Information "User [$($csOnlineUser.userPrincipalName)] has a Teams Phone (Phone System/MCOEV) license"
    }

    if ($null -ne $correlatedAccount) {
        $correlatedAccount.EnterpriseVoiceEnabled = if ($correlatedAccount.EnterpriseVoiceEnabled) { 'True' } else { 'False' }
        $outputContext.PreviousData = $correlatedAccount.PSObject.Copy() # Required to disconnect the object from $correlatedAccount
        $outputContext.PreviousData.CallingLineIdentity = $correlatedAccount.CallingLineIdentity.Name

        $actionMessage = "Validating DepartmentId"
        if ([string]::IsNullOrEmpty($actionContext.Data.DepartmentId)) {
            throw "Department id is empty"
        }

        # Lookup CallingLineIdentity where description is equal to $actionContext.Data.DepartmentId
        $actionMessage = "Looking up CallingLineIdentity for department id [$($actionContext.Data.DepartmentId)]"
        $callingLineIdentity = Get-CsCallingLineIdentity | Where-Object { $_.Description -eq $actionContext.Data.DepartmentId }
        if (@($callingLineIdentity).Count -eq 0) {
            throw "No CallingLineIdentity found for department id [$($actionContext.Data.DepartmentId)]"
        }
        elseif (@($callingLineIdentity).Count -gt 1) {
            throw "Multiple CallingLineIdentity found for department id [$($actionContext.Data.DepartmentId)]"
        }

        # Check if CallingLineIdentity needs to be updated
        $currentCallingLineIdentity = if ($null -ne $correlatedAccount.CallingLineIdentity) { "Tag:$($correlatedAccount.CallingLineIdentity)" } else { $null }
        $changeCallingLineIdentityName = ($currentCallingLineIdentity -ne $callingLineIdentity.Identity)

        # Check if EnterpriseVoiceEnabled needs to be updated (and exists in the mapping - could be improved by using get-member, but the current method is sufficient)
        $changeEnterpriseVoiceEnabled = ($correlatedAccount.EnterpriseVoiceEnabled -ne $actionContext.Data.EnterpriseVoiceEnabled -and $null -ne $actionContext.Data.EnterpriseVoiceEnabled)

        if ($changeCallingLineIdentityName -or $changeEnterpriseVoiceEnabled) {
            $action = 'UpdateAccount'
        }
        else {
            $action = 'NoChanges'
        }
        $outputContext.Data = $correlatedAccount
    }
    else {
        $action = 'NotFound'
    }

    # Process
    switch ($action) {
        'UpdateAccount' {
            if ($changeCallingLineIdentityName) {
                $callingLineIdentityName = $callingLineIdentity.Identity.Substring(4) # Remove 'Tag:' prefix

                $actionMessage = "Updating CallingLineIdentity to [$callingLineIdentityName] for [$($correlatedAccount.userPrincipalName)]"
                $auditMessage = "CallingLineIdentity updated to [$callingLineIdentityName] for [$($correlatedAccount.userPrincipalName)]"

                if (-not($actionContext.DryRun -eq $true)) {
                    Write-Information $actionMessage
                    $null = Grant-CsCallingLineIdentity -Identity $correlatedAccount.Identity -PolicyName $callingLineIdentity.Identity -ErrorAction Stop
                }
                else {
                    $auditMessage = "[DryRun] $auditMessage"
                    Write-Information $auditMessage
                }

                $outputContext.Data.CallingLineIdentity = $callingLineIdentityName
                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = $action
                        Message = $auditMessage
                        IsError = $false
                    })
            }

            if ($changeEnterpriseVoiceEnabled) {
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
                $outputContext.Data.EnterpriseVoiceEnabled = $actionContext.Data.EnterpriseVoiceEnabled
                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = $action
                        Message = $auditMessage
                        IsError = $false
                    })
            }

            $outputContext.AccountReference = $correlatedAccount.Identity
            $outputContext.Success = $true
            break
        }

        'NoChanges' {
            Write-Information "No changes required to MicrosoftTeams-Voip account [$($correlatedAccount.userPrincipalName)]"
            $outputContext.Success = $true
            break
        }

        'NotFound' {
            Write-Information "MicrosoftTeams-Voip account: [$($actionContext.References.Account)] could not be found, indicating that it may have been deleted"
            $outputContext.Success = $false
            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Message = "MicrosoftTeams-Voip account: [$($actionContext.References.Account)] could not be found, indicating that it may have been deleted"
                    IsError = $true
                })
            break
        }
    }
}
catch {
    $outputContext.Success = $false
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