#-------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation.  All rights reserved.
#-------------------------------------------------------------------------

#
# Returns given date-time by locale 'en-us'
#
Function Global:GetDateTimeLocaleEnUs
{
    param
    (
        [DateTime]
        [parameter(mandatory=$true)]
        $DateTime
    )

    $culture = New-Object System.Globalization.CultureInfo 'en-us'
    $dateTime = $($DateTime.ToUniversalTime()).ToString($culture)

    Write-Output $dateTime
}

#
# Emit events in order to track usage and problems through ADHealth
#
Function Global:WriteEventLog
{
    param
    (
        [int]
        [parameter(mandatory=$true)]
        $EventId,

        [string]
        [parameter(mandatory=$true)]
        $Message
    )

    Write-EventLog -LogName "Application" -Source "Directory Synchronization" -EventID $EventId -EntryType Information -Message $Message -Category 0
}

Function Global:GetValidADForestCredentials
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $ADForestName
    )

    while ($true)
    {
        $ADForestCredential = Get-Credential -Message "Please enter credentials for an AD account that has permissions to read all attributes on the target object. Use the fully qualified domain name of the account (Example: CONTOSO\admin):"
        $networkCredential = $ADForestCredential.GetNetworkCredential()

        if ($networkCredential.UserName.Contains("@") -or $networkCredential.Domain.ToString() -eq "")
        {
            "Credential should use the fully qualified domain name of the account. Example: CONTOSO\admin" | Write-Host -fore Red
            Write-Host "`r`n"
        } else {
            try
            {
                $credentialTestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext([System.DirectoryServices.ActiveDirectory.DirectoryContextType]::Forest, $ADForestName, $ADForestCredential.UserName, $networkCredential.Password)
                $credentialTestForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($credentialTestContext)
                return $ADForestCredential
            }
            catch
            {
                "Invalid Credentials. Details: $($_.Exception.Message)" | Write-Host -fore Red
                Write-Host "`r`n"
            }
        }

        $retryCredentialOptions = [System.Management.Automation.Host.ChoiceDescription[]] @("&Provide credentials", "&Back to previous")
        $retryCredential = !($host.UI.PromptForChoice("Credentials Invalid", "Provide credentials again or return to previous menu?", $retryCredentialOptions, 0))

        if (!$retryCredential)
        {
            return $null
        }
    }
}

Function Global:GetADConnectors
{
    $adConnectors = Get-ADSyncConnector | Where-Object {$_.ConnectorTypeName -eq "AD"}
    Write-Output $adConnectors
}

Function Global:GetADConnectorByName
{
    param
    (
        [string]
        [parameter(mandatory=$false)]
        [AllowNull()]
        $ADConnectorName,

        [string]
        [parameter(mandatory=$false)]
        $PromptMessage
    )

    while ([string]::IsNullOrEmpty($ADConnectorName))
    {
        $adConnectors = GetADConnectors
        if ($adConnectors -eq $null)
        {
            "No AD Connector is found." | ReportError
            Write-Host "`r`n"
            $ADConnectorName = $null
            return
        }

        Write-Host "`r`n"

        Write-Host "List of AD Connectors:"
        Write-Host "----------------------"

        foreach ($adConnector in $adConnectors)
        {
            Write-Host $adConnector.Name
        }

        if ($adConnectors.length -eq 1)
        {
            $ADConnectorName = $adConnectors[0].Name
        }
        else
        {
            Write-Host "`r`n"

            $ADConnectorName = Read-Host $PromptMessage
        }
    }

    $adConnector = Get-ADSyncConnector | Where-Object {($_.ConnectorTypeName -eq "AD") -and ($_.Name -eq $ADConnectorName)}
    Write-Output $adConnector
}

Function Global:GetAADConnector
{
   $aadConnectors = Get-ADSyncConnector | Where-Object {$_.Identifier -eq "b891884f-051e-4a83-95af-2544101c9083"}

   if ($aadConnectors -eq $null)
   {
        Write-Output $null
   }
   else
   {
        Write-Output $aadConnectors[0]
   }
}

Function Global:GetAADTenantName
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector]
        [parameter(mandatory=$true)]
        $AADConnector
    )

    $aadConnectorUserName = $AADConnector.ConnectivityParameters["UserName"].Value
    $aadTenantName = $($aadConnectorUserName.Split('@'))[1]

    Write-Output $aadTenantName
}

Function Global:GetCSObject
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $ConnectorName,

        [string]
        [parameter(mandatory=$true)]
        $DistinguishedName
    )

    Try
    {
        $csObject = Get-ADSyncCSObject -ConnectorName $ConnectorName -DistinguishedName $DistinguishedName
        Write-Output $csObject
    }
    Catch
    {
        Write-Output $null
    }
}

Function Global:GetCSObjectByIdentifier
{
    param
    (
        [Guid]
        [parameter(mandatory=$true)]
        $CsObjectId
    )

    Try
    {
        $csObject = Get-ADSyncCSObject -Identifier $CsObjectId
        Write-Output $csObject
    }
    Catch
    {
        Write-Output $null
    }
}

Function Global:GetMVObjectByIdentifier
{
    param
    (
        [Guid]
        [parameter(mandatory=$true)]
        $MvObjectId
    )

    Try
    {
        $mvObject = Get-ADSyncMVObject -Identifier $MvObjectId
        Write-Output $mvObject
    }
    Catch
    {
        Write-Output $null
    }
}

Function Global:GetTargetCSObjectId
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        $MvObject,

        [Guid]
        [parameter(mandatory=$true)]
        $TargetConnectorId
    )

    foreach ($mvObjectLink in $MvObject.Lineage)
    {
        if ($mvObjectLink.ConnectorId -eq $TargetConnectorId)
        {
            Write-Output $mvObjectLink.ConnectedCsObjectId
            return
        }
    }

    Write-Output $null
}

Function Global:IsStagingModeEnabled
{
    $isStagingModeEnabled = $false
    $globalParameters = Get-ADSyncGlobalSettingsParameter
    $stagingModeGlobalParameter = $globalParameters | Where-Object {$_.Name -eq "Microsoft.Synchronize.StagingMode"}

    if ($stagingModeGlobalParameter -ne $null)
    {
        $isStagingModeEnabled = $stagingModeGlobalParameter.Value
    }

    Write-Output $isStagingModeEnabled
}

Function Global:ConvertADObjectToHashTable
{
    param
    (
        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        $AdObject,

        [bool]
        [parameter(mandatory=$false)]
        $AllAttributes
    )

    $ADObjectHashTable = @{}

    if ($AllAttributes)
    {
        foreach ($attributeName in $AdObject.Keys)
        {
            $attributeValues = [System.Collections.ArrayList] $AdObject[$attributeName]

            $ADObjectHashTable[$attributeName] = New-Object System.Collections.Generic.List[String]

            foreach ($attributeValue in $attributeValues)
            {
                $ADObjectHashTable[$attributeName].Add($attributeValue)
            }
        }
    }
    else
    {
        foreach ($attributeName in $global:ADObjectAttributes)
        {
            if (!$AdObject.ContainsKey($attributeName))
            {
                continue
            }

            $attributeValues = [System.Collections.ArrayList] $AdObject[$attributeName]

            $ADObjectHashTable[$attributeName] = New-Object System.Collections.Generic.List[String]

            foreach ($attributeValue in $attributeValues)
            {
                $ADObjectHashTable[$attributeName].Add($attributeValue)
            }
        }
    }

    Write-Output $ADObjectHashTable
}

Function Global:ConvertCSObjectToHashTable
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        $CsObject,

        [string[]]
        [parameter(mandatory=$true)]
        $CsObjectAttributes
    )

    $CsObjectHashTable = @{}

    foreach ($attributeName in $CsObjectAttributes)
    {
        if (!$CsObject.Attributes.Contains($attributeName))
        {
            continue
        }

        $attributeValues = $CsObject.Attributes[$attributeName].Values

        $CsObjectHashTable[$attributeName] = New-Object System.Collections.Generic.List[String]

        foreach ($attributeValue in $attributeValues)
        {
            $CsObjectHashTable[$attributeName].Add($attributeValue)
        }
    }

    Write-Output $CsObjectHashTable
}

Function Global:ConvertMVObjectToHashTable
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        $MvObject
    )

    $MvObjectHashTable = @{}

    foreach ($attributeName in $global:MvObjectAttributes)
    {
        if (!$MvObject.Attributes.Contains($attributeName))   
        {
            continue
        }

        $attributeValues = $MvObject.Attributes[$attributeName].Values

        $MvObjectHashTable[$attributeName] = New-Object System.Collections.Generic.List[String]

        foreach ($attributeValue in $attributeValues)
        {
            $MvObjectHashTable[$attributeName].Add($attributeValue)
        }
    }

    Write-Output $MvObjectHashTable
}

Function Global:ConvertAADUserObjectToHashTable
{
    param
    (
        [Microsoft.Online.Administration.User]
        [parameter(mandatory=$true)]
        $AadUserObject
    )

    $AadUserObjectHashTable = @{}

    if ($AadUserObject.DisplayName)
    {
        $AadUserObjectHashTable["DisplayName"] = New-Object System.Collections.Generic.List[String]

        $AadUserObjectHashTable["DisplayName"].Add($AadUserObject.DisplayName)
    }

    if ($AadUserObject.ImmutableId)
    {
        $AadUserObjectHashTable["ImmutableId"] = New-Object System.Collections.Generic.List[String]

        $AadUserObjectHashTable["ImmutableId"].Add($AadUserObject.ImmutableId)
    }

    if ($AadUserObject.IsLicensed)
    {
        $AadUserObjectHashTable["IsLicensed"] = New-Object System.Collections.Generic.List[String]
    
        $AadUserObjectHashTable["IsLicensed"].Add($AadUserObject.IsLicensed)    
    }
    
    if ($AadUserObject.LastDirSyncTime)
    {
        $AadUserObjectHashTable["LastDirSyncTime"] = New-Object System.Collections.Generic.List[String]
    
        $AadUserObjectHashTable["LastDirSyncTime"].Add($AadUserObject.LastDirSyncTime)
    }

    if ($AadUserObject.UserPrincipalName)
    {
        $AadUserObjectHashTable["UserPrincipalName"] = New-Object System.Collections.Generic.List[String]

        $AadUserObjectHashTable["UserPrincipalName"].Add($AadUserObject.UserPrincipalName)    
    }
    
    if ($AadUserObject.ProxyAddresses)
    {
        $AadUserObjectHashTable["ProxyAddresses"] = New-Object System.Collections.Generic.List[String]
    
        foreach ($proxyAddress in $AadUserObject.ProxyAddresses)
        {
            $AadUserObjectHashTable["ProxyAddresses"].Add($proxyAddress)    
        }
    }

    Write-Output $AadUserObjectHashTable
}

Function Global:ConvertAADContactObjectToHashTable
{
    param
    (
        [Microsoft.Online.Administration.Contact]
        [parameter(mandatory=$true)]
        $AadContactObject
    )

    $AadContactObjectHashTable = @{}

    if ($AadContactObject.DisplayName)
    {
        $AadContactObjectHashTable["DisplayName"] = New-Object System.Collections.Generic.List[String]

        $AadContactObjectHashTable["DisplayName"].Add($AadContactObject.DisplayName)
    }

    if ($AadContactObject.ImmutableId)
    {
        $AadContactObjectHashTable["ImmutableId"] = New-Object System.Collections.Generic.List[String]

        $AadContactObjectHashTable["ImmutableId"].Add($AadContactObject.ImmutableId)
    }
    
    if ($AadContactObject.LastDirSyncTime)
    {
        $AadContactObjectHashTable["LastDirSyncTime"] = New-Object System.Collections.Generic.List[String]
    
        $AadContactObjectHashTable["LastDirSyncTime"].Add($AadContactObject.LastDirSyncTime)
    }

    if ($AadContactObject.UserPrincipalName)
    {
        $AadContactObjectHashTable["UserPrincipalName"] = New-Object System.Collections.Generic.List[String]

        $AadContactObjectHashTable["UserPrincipalName"].Add($AadContactObject.UserPrincipalName)    
    }
    
    if ($AadContactObject.ProxyAddresses)
    {
        $AadContactObjectHashTable["ProxyAddresses"] = New-Object System.Collections.Generic.List[String]
    
        foreach ($proxyAddress in $AadContactObject.ProxyAddresses)
        {
            $AadContactObjectHashTable["ProxyAddresses"].Add($proxyAddress)    
        }
    }

    Write-Output $AadContactObjectHashTable
}

Function Global:ConvertAADGroupObjectToHashTable
{
    param
    (
        [Microsoft.Online.Administration.Group]
        [parameter(mandatory=$true)]
        $AadGroupObject
    )

    $AadGroupObjectHashTable = @{}

    if ($AadGroupObject.DisplayName)
    {
        $AadGroupObjectHashTable["DisplayName"] = New-Object System.Collections.Generic.List[String]

        $AadGroupObjectHashTable["DisplayName"].Add($AadGroupObject.DisplayName)
    }

    if ($AadGroupObject.ImmutableId)
    {
        $AadGroupObjectHashTable["ImmutableId"] = New-Object System.Collections.Generic.List[String]

        $AadGroupObjectHashTable["ImmutableId"].Add($AadGroupObject.ImmutableId)
    }
    
    if ($AadGroupObject.LastDirSyncTime)
    {
        $AadGroupObjectHashTable["LastDirSyncTime"] = New-Object System.Collections.Generic.List[String]
    
        $AadGroupObjectHashTable["LastDirSyncTime"].Add($AadGroupObject.LastDirSyncTime)
    }
    
    if ($AadGroupObject.ProxyAddresses)
    {
        $AadGroupObjectHashTable["ProxyAddresses"] = New-Object System.Collections.Generic.List[String]
    
        foreach ($proxyAddress in $AadGroupObject.ProxyAddresses)
        {
            $AadGroupObjectHashTable["ProxyAddresses"].Add($proxyAddress)    
        }
    }

    Write-Output $AadGroupObjectHashTable
}


Function ReportOutput
{
    param
    (
        [string]
        [parameter(mandatory=$false, ValueFromPipeline=$True)]
        $Output,

        [string]
        [parameter(mandatory=$false)]
        $PropertyName = [string]::Empty,

        [string]
        [parameter(mandatory=$false)]
        $PropertyValue = [string]::Empty
    )

    if ($isNonInteractiveMode)
    {
        Write-AscOutput -Message $Output -PropertyName $PropertyName -PropertyValue $PropertyValue -OutputType 'Output'
    }
    else
    {
        Write-Host $Output -fore Green
    }
}

Function ReportError
{
    param
    (
        [string]
        [parameter(mandatory=$true, ValueFromPipeline=$True)]
        $ErrorString,

        [string]
        [parameter(mandatory=$false)]
        $PropertyName = [string]::Empty,

        [string]
        [parameter(mandatory=$false)]
        $PropertyValue = [string]::Empty

    )

    if ($isNonInteractiveMode)
    {
        Write-AscOutput -Message $ErrorString -PropertyName $PropertyName -PropertyValue $PropertyValue -OutputType 'Error'
    }
    else
    {
        Write-Host $ErrorString -fore Red
    }
}

Function ReportWarning
{
    param
    (
        [string]
        [parameter(mandatory=$true, ValueFromPipeline=$True)]
        $WarningString,

        [string]
        [parameter(mandatory=$false)]
        $PropertyName = [string]::Empty,

        [string]
        [parameter(mandatory=$false)]
        $PropertyValue = [string]::Empty
    )

    if ($isNonInteractiveMode)
    {
        Write-AscOutput -Message $WarningString -PropertyName $PropertyName -PropertyValue $PropertyValue -OutputType 'Warning'   
    }
    else
    {
        Write-Host $WarningString -fore Cyan
    }
}

Function Write-AscOutput
{
    param
    (
        [string]
        [parameter(mandatory=$false, ValueFromPipeline=$True)]
        $Message = [string]::Empty,

        [string]
        [parameter(mandatory=$false)]
        $PropertyName = [string]::Empty,

        [string]
        [parameter(mandatory=$false)]
        $PropertyValue = [string]::Empty,

        [string]
        [parameter(mandatory=$true)]
        $OutputType
    )

    $obj = New-Object PSCustomObject -Property @{
        $AscCustomScriptObjectProperty      = $true          
        PropertyName                        = $PropertyName                 
        PropertyValue                       = $PropertyValue              
        Message                             = $Message
        OutputType                          = $OutputType            
    }

    Write-Output $obj  
}
# SIG # Begin signature block
# MIIoNgYJKoZIhvcNAQcCoIIoJzCCKCMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAj37X/gCZTRN9G
# Coga+KauNaiZYyma0dopRtmVHVrXJ6CCDYIwggYAMIID6KADAgECAhMzAAADXJXz
# SFtKBGrPAAAAAANcMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwNDA2MTgyOTIyWhcNMjQwNDAyMTgyOTIyWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDijA1UCC84R0x+9Vr/vQhPNbfvIOBFfymE+kuP+nho3ixnjyv6vdnUpgmm6RT/
# pL9cXL27zmgVMw7ivmLjR5dIm6qlovdrc5QRrkewnuQHnvhVnLm+pLyIiWp6Tow3
# ZrkoiVdip47m+pOBYlw/vrkb8Pju4XdA48U8okWmqTId2CbZTd8yZbwdHb8lPviE
# NMKzQ2bAjytWVEp3y74xc8E4P6hdBRynKGF6vvS6sGB9tBrvu4n9mn7M99rp//7k
# ku5t/q3bbMjg/6L6mDePok6Ipb22+9Fzpq5sy+CkJmvCNGPo9U8fA152JPrt14uJ
# ffVvbY5i9jrGQTfV+UAQ8ncPAgMBAAGjggF/MIIBezArBgNVHSUEJDAiBgorBgEE
# AYI3TBMBBgorBgEEAYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUXgIsrR+tkOQ8
# 10ekOnvvfQDgTHAwRQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEWMBQGA1UEBRMNMjMzMTEwKzUwMDg2ODAfBgNVHSMEGDAWgBRI
# bmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEt
# MDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBABIm
# T2UTYlls5t6i5kWaqI7sEfIKgNquF8Ex9yMEz+QMmc2FjaIF/HQQdpJZaEtDM1Xm
# 07VD4JvNJEplZ91A4SIxjHzqgLegfkyc384P7Nn+SJL3XK2FK+VAFxdvZNXcrkt2
# WoAtKo0PclJOmHheHImWSqfCxRispYkKT9w7J/84fidQxSj83NPqoCfUmcy3bWKY
# jRZ6PPDXlXERRvl825dXOfmCKGYJXHKyOEcU8/6djs7TDyK0eH9ss4G9mjPnVZzq
# Gi/qxxtbddZtkREDd0Acdj947/BTwsYLuQPz7SNNUAmlZOvWALPU7OOVQlEZzO8u
# Ec+QH24nep/yhKvFYp4sHtxUKm1ZPV4xdArhzxJGo48Be74kxL7q2AlTyValLV98
# u3FY07rNo4Xg9PMHC6sEAb0tSplojOHFtGtNb0r+sioSttvd8IyaMSfCPwhUxp+B
# Td0exzQ1KnRSBOZpxZ8h0HmOlMJOInwFqrCvn5IjrSdjxKa/PzOTFPIYAfMZ4hJn
# uKu15EUuv/f0Tmgrlfw+cC0HCz/5WnpWiFso2IPHZyfdbbOXO2EZ9gzB1wmNkbBz
# hj8hFyImnycY+94Eo2GLavVTtgBiCcG1ILyQabKDbL7Vh/OearAxcRAmcuVAha07
# WiQx2aLghOSaZzKFOx44LmwUxRuaJ4vO/PRZ7EzAMIIHejCCBWKgAwIBAgIKYQ6Q
# 0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNh
# dGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5
# WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQD
# Ex9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4
# BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe
# 0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato
# 88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v
# ++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDst
# rjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN
# 91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4ji
# JV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmh
# D+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbi
# wZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8Hh
# hUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaI
# jAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTl
# UAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNV
# HQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQF
# TuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29m
# dC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNf
# MjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNf
# MjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcC
# ARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnlj
# cHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5
# AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oal
# mOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0ep
# o/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1
# HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtY
# SWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInW
# H8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZ
# iWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMd
# YzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7f
# QccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKf
# enoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOpp
# O6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZO
# SEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4xCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jv
# c29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAANclfNIW0oEas8AAAAAA1ww
# DQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJa7XmiD
# 0MBcuOi5ImObF7o+imkYAuGPoxYS/eXb9JOWMEIGCisGAQQBgjcCAQwxNDAyoBSA
# EgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20w
# DQYJKoZIhvcNAQEBBQAEggEAWZFJ6e4YyImdN3j1WYYPr/tOK6T1YWa50VqiYwW/
# 9qm1EfEB3LXbqpEachVEX3S3l63Fno2T/QeaXcuMOFXtczHq0k1MCSZs00u64txR
# jw8ny50o+1FoYRa0w+ioHtH8sX5nY0lmi+G8FoC+DjbO9OclJHgAeCChbCLL4A/+
# G/+Aej3MW6yS6+mld5L69J9+erjgm4zaK3XHu4+BeRBeBwKSZ30VPKLSeF8I0PS4
# P2IzV+TsDuMiAUFqXwHoO5/di3NlynjsCMWlSLNIPLDGExKVqUqCDOYnURFzKz40
# 65BcQmc71g++Xf2AuCfXUT5LsXosDmU50u/IrfmfGzuEX6GCF5QwgheQBgorBgEE
# AYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUD
# BAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoD
# ATAxMA0GCWCGSAFlAwQCAQUABCB/eJFSidn8hc7HCbj2CpDV7l8KODz0xTi/rpLt
# CJlPOAIGZQQ0jj13GBMyMDIzMTAwNDE5Mjg1NC42MjVaMASAAgH0oIHRpIHOMIHL
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxN
# aWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRT
# UyBFU046RjAwMi0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAc4PGPdFl+fG/wABAAAB
# zjANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAe
# Fw0yMzA1MjUxOTEyMDhaFw0yNDAyMDExOTEyMDhaMIHLMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmlj
# YSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046RjAwMi0wNUUw
# LUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC5CkwZ1yjYx3fnKTw/VnzwGGhK
# OIjqMDSuHdGg8JoJ2LN2nBUUkAwxhYAR4ZQWg9QbjxZ/DWrD2xeUwLnKOKNDNthX
# 9vaKj+X5Ctxi6ioTVU7UB5oQ4wGpkV2kmfnp0RYGdhtc58AaoUZFcvhdBlJ2yETw
# uCuEV6pk4J7ghGymszr0HVqR9B2MJjV8rePL+HGIzIbYLrk0jWmaKRRPsFfxKKw3
# njFgFlSqaBA4SVuV0FYE/4t0Z9UjXUPLqw+iDeLUv3sp3h9M4oNIZ216VPlVlf3F
# OFRLlZg8eCeX4xlaBjWia95nXlXMXQWqaIwkgN4TsRzymgeWuVzMpRPBWk6gOjzx
# wXnjIcWqx1lPznISv/xtn1HpB+CIF5SPKkCf8lCPiZ1EtB01FzHRj+YhRWJjsRl1
# gLW1i0ELrrWVAFrDPrIshBKoz6SUAyKD7yPx649SyLGBo/vJHxZgMlYirckf9ekl
# prNDeoslhreIYzAJrMJ+YoWn9Dxmg/7hGC/XH8eljmJqBLqyHCmdgS+WArj84ciR
# GsmqRaUB/4hFGUkLv1Ga2vEPtVByUmjHcAppJR1POmi1ATV9FusOQQxkD2nXWSKW
# fKApD7tGfNZMRvkufHFwGf5NnN0Aim0ljBg1O5gs43Fok/uSe12zQL0hSP9Jf+iC
# L+NPTPAPJPEsbdYavQIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFDD7CEZAo5MMjpl+
# FWTsUyn54oXFMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1Ud
# HwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3Js
# L01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggr
# BgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIw
# MTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgw
# DgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCXIBYW/0UVTDDZO/fQ
# 2XstNC4DZG8RPbrlZHyFt57z/VWqPut6rugayGW1UcvJuxf8REtiTtmf5SQ5N2pu
# 0nTl6O4BtScIvM/K8pe/yj77x8u6vfk8Q6SDOZoFpIpVkFH3y67isf4/SfoN9M2n
# Lb93po/OtlM9AcWTJbqunzC+kmeLcxJmCxLcsiBMJ6ZTvSNWQnicgMuv7PF0ip9H
# YjzFWoNq8qnrs7g++YGPXU7epl1KSBTr9UR7Hn/kNcqCiZf22DhoZPVP7+vZHTY+
# OXoxoEEOnzAbAlBCup/wbXNJissiK8ZyRJXT/R4FVmE22CSvpu+p5MeRlBT42pkI
# hhMlqXlsdQdT9cWItiW8yWRpaE1ZI1my9FW8JM9DtCQti3ZuGHSNpvm4QAY/61ry
# rKol4RLf5F+SAl4ozVvM8PKMeRdEmo2wOzZK4ME7D7iHzLcYp5ucw0kgsy396fac
# zsXdnLSomXMArstGkHvt/F3hq2eESQ2PgrX+gpdBo8uPV16ywmnpAwYqMdZM+yH6
# B//4MsXEu3Rg5QOoOWdjNVB7Qm6MPJg+vDX59XvMmibAzbplxIyp7S1ky7L+g3hq
# 6KxlKQ9abUjYpaOFnHtKDFJ+vxzncEMVEV3IHQdjC7urqOBgO7vypeIwjQ689qu2
# NNuIQ6cZZgMn8EvSSWRwDG8giTCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkA
# AAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRl
# IEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVow
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX
# 9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1q
# UoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8d
# q6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byN
# pOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2k
# rnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4d
# Pf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgS
# Uei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8
# QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6Cm
# gyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzF
# ER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQID
# AQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU
# KqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1
# GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0
# dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0
# bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMA
# QTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbL
# j+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1p
# Y3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0w
# Ni0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIz
# LmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwU
# tj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN
# 3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU
# 5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5
# KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGy
# qVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB6
# 2FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltE
# AY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFp
# AUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcd
# FYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRb
# atGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQd
# VTNYs6FwZvKhggNNMIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2Eg
# T3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkYwMDItMDVFMC1E
# OTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEw
# BwYFKw4DAhoDFQBdjZUbFNAyCkVE6DdVWyizTYQHzKCBgzCBgKR+MHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6Me+5zAiGA8y
# MDIzMTAwNDEwMzcyN1oYDzIwMjMxMDA1MTAzNzI3WjB0MDoGCisGAQQBhFkKBAEx
# LDAqMAoCBQDox77nAgEAMAcCAQACAg8/MAcCAQACAhN5MAoCBQDoyRBnAgEAMDYG
# CisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEA
# AgMBhqAwDQYJKoZIhvcNAQELBQADggEBAFjod1nuxAVmunl6IriYY0WpSrSx/vzz
# 9BDbUYwWY4yEFLg+jUwemQCt2boTFjcQDww7mCHsP9NfI58i4pMcZO9Nsa/mMTSz
# qiWlUU3FyOYPoF8BwN3Ioq4/pkDe6fX8erhDdg15PhBogRkmI501gFMrVTL0ODG2
# of/Psb9DTni/bmA0aSUuPOktc/RY8SUwyCXhBiSqqBjyiYQVf8zrnmvayGAz0smI
# 0KRyk8f+CWg0eUrq5kOmYd5HicZXMZnzoPEsnjYwPxkh1JrngGNwMAXy3+XOtq4R
# PUmL+OcZ9aTboJFhEaTTWLI/foDDsbUFhY3a9YFKsXY5wJ14KywvzwgxggQNMIIE
# CQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYw
# JAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAc4PGPdF
# l+fG/wABAAABzjANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqG
# SIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCBej+18zIeWglBTBT959IkhySw/+ugp
# 1MrV42iYo/RdQjCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIDJsz0F6L0XD
# Um53JRBfNMZszKsllLDMiaFZ3PL/LqxnMIGYMIGApH4wfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAHODxj3RZfnxv8AAQAAAc4wIgQg1+caYgNLKhNy
# gXcmWt0L0jo/QJsgWkeqnT9LPZdM47kwDQYJKoZIhvcNAQELBQAEggIABAmlfkmy
# xI6FgGy0vEDJVOIMRuyKzMzhIRzAfWtUA8AZbxsuMJKzIUpykW2dcOqzKX6nX2rx
# ZwT3IXs17bzyOtV15+fVNnm2vRw6MEV9tywVKpsRKPJoQN78G1d5MIajnIo0x9ZE
# eRE94pYBRbo2VCHrjQ62Cnmhooj9lE9PXLT2i5hQmmtWmZzWQAgOmT0hawsSjYu7
# Uhz7I6RidKXryNDFLkYJS0ZyF12ghIkYIjbFpFCgpPH2DWN8zI9u52av7sBwANu+
# UpNXvielvIUlSPu+mrN8nKENVZLtESXyLUuQ/bvE9jR6bqf7bG+GHaSO+As0gvns
# ypfoVsgsC6iw0bF5j4myufdg/fuv5Ai2ZfJ6IEmkzfYUJmEZOECXhB5fNCNxMOKL
# 1Ny6RpKCuAnU9QsuwQqCMgeH3HRNodviEmm/TnyWkKK8YwJJ9HEESDTgG6wweyBu
# iyb/3cinumpdQ/d3EJmKPPK8IahBjiZ7IjxSWfi9t5ec1WV2eeDUTtSRVy8Fy5w6
# oxSBOwlP213ISQpaPOAfe9A/pqnMhqy3a24QSrxt5v2O4woBOZIFWnQLXLfuP6Ng
# 1KXqKdYREnWzPxAYLYelo2Q9luUuaQtIDSQ6E7/lofs19cc5Y6aAfb1V7nsSLuHe
# M8nd0G/b1w3Zl0j8ITAO2DV9+rETXKNnU+E=
# SIG # End signature block
