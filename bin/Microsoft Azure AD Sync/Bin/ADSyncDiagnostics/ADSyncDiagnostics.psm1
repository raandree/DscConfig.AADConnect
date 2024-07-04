Import-Module ADSync

Get-ChildItem -Path "$PSScriptRoot\PSScripts\*.ps1" | % { . $_.FullName }

Set-Variable ADSyncDiagnosticsOutputPath "$env:ALLUSERSPROFILE\AADConnect" -option Constant

$global:AADTenantCredential = $null
$global:AscCustomScriptObjectProperty = 'IsAscCustomScriptObject'

#
# AADConnect General Diagnostics
#
#    Invoke-ADSyncDiagnostics
#
#
# Complete Password Hash Sync Diagnostics
#
#    Invoke-ADSyncDiagnostics -PasswordSync
#
#
# Password Hash Sync Single Object Diagnostics
#
#    Invoke-ADSyncDiagnostics -PasswordSync -ADConnectorName <AD Connector Name> -DistinguishedName <Distinguished Name>
#
function Invoke-ADSyncDiagnostics
{
    param
    (
        [switch]
        [parameter(parametersetname="passwordsync", mandatory=$false)]
        $PasswordSync,

        [string]
        [parameter(parametersetname="passwordsync", mandatory=$false)]
        $ADConnectorName,

        [string]
        [parameter(parametersetname="passwordsync", mandatory=$false)]
        $DistinguishedName,

        [string]
        [parameter(parametersetname="passwordsync", mandatory=$false)]
        $MemberADConnectorName,

        [string]
        [parameter(parametersetname="passwordsync", mandatory=$false)]
        $MemberDistinguishedName,

        [switch]
        [parameter(parametersetname="passwordsync", mandatory=$false)]
        $NonInteractiveMode
    )

    # Currently Non-interactive mode is used to run script from ASC.
    $isNonInteractiveMode = $NonInteractiveMode.IsPresent

    if ($PasswordSync -and ([string]::IsNullOrEmpty($ADConnectorName) -or [string]::IsNullOrEmpty($DistinguishedName)))
    {
        if ($isNonInteractiveMode)
        {
            if ([string]::IsNullOrEmpty($ADConnectorName))
            {
                "ADConnectorName is a mandatory parameter in Non-Interactive Mode. Either provide AdConnectorName or rerun the script without Non-Interactive mode." | Write-Host -fore Red
                return
            }

            $result = DiagnosePasswordHashSyncNonInteractiveMode -ADConnectorName $ADConnectorName

            if ($result)
            {
                # In Non-Interactive mode just output the objects which are logged for ASC purpose.
                $result | Where-Object {(Get-Member -InputObject $_ -Name  $AscCustomScriptObjectProperty) -ne $null}	
            }
        }
        else
        {
            # Complete Password Hash Sync Diagnostics
            DiagnosePasswordHashSync
            PromptPasswordSyncSingleObjectDiagnostics
        }
    }
    elseif ($PasswordSync)
    {
        if ($isNonInteractiveMode)
        {
            $result = DiagnosePasswordHashSyncNonInteractiveMode -ADConnectorName $ADConnectorName
            $result += DiagnosePasswordSyncSingleObject -ADConnectorName $ADConnectorName -DistinguishedName $DistinguishedName

            if ($result)
            {
                # In Non-Interactive mode just output the objects which are logged for ASC purpose.
                $result | Where-Object {(Get-Member -InputObject $_ -Name  $AscCustomScriptObjectProperty) -ne $null}	
            }
        }
        else
        {
            # Password Hash Sync Single Object Diagnostics
            DiagnosePasswordSyncSingleObject -ADConnectorName $ADConnectorName -DistinguishedName $DistinguishedName
        }
    }
    else
    {
        if ($isNonInteractiveMode)
        {
            if ([string]::IsNullOrEmpty($ADConnectorName) -or [string]::IsNullOrEmpty($DistinguishedName))
            {
                "ADConnectorName and DistinguishedName are mandatory parameters in Non-Interactive Mode. Either provide all the mandatory parameters or rerun the script without Non-Interactive mode." | Write-Host -fore Red
                return
            }
            elseif ([string]::IsNullOrEmpty($MemberADConnectorName) -or [string]::IsNullOrEmpty($MemberDistinguishedName))
            {
                $result = Debug-ADSyncObjectSynchronizationIssuesNonInteractiveMode -ADConnectorName $ADConnectorName -ObjectDN $DistinguishedName -DiagnosticOption $global:DiagnoseObjectSyncIssues

                if ($result)
                {
                    # In Non-Interactive mode just output the objects which are logged for ASC purpose.
                    $result | Where-Object {(Get-Member -InputObject $_ -Name  $AscCustomScriptObjectProperty) -ne $null}	
                }
            }
            else
            {
                $result = Debug-ADSyncGroupMembershipSynchronizationIssuesNonInteractiveMode -GroupADConnectorName $ADConnectorName -GroupDN $DistinguishedName -MemberADConnectorName $MemberADConnectorName -MemberDN $MemberDistinguishedName

                if ($result)
                {
                    # In Non-Interactive mode just output the objects which are logged for ASC purpose.
                    $result | Where-Object {(Get-Member -InputObject $_ -Name  $AscCustomScriptObjectProperty) -ne $null}	
                }
            }
        }
        else
        {
            MainMenu
        }

        return
    }

    Write-Host "`r`n"
    "For more help:" | Write-Host -fore Yellow
    "+ Please see - https://go.microsoft.com/fwlink/?linkid=847231 or" | Write-Host -fore Yellow
    "+ Open a service request through Azure Portal or Office 365 Admin Portal." | Write-Host -fore Yellow
    Write-Host "`r`n"
}

#
# Main Troubleshooting Menu
#
function MainMenu
{
    cls
    $isQuit = $false

    while ($true)
    {
        Show-MainMenu

        $selection = Read-Host "`tPlease make a selection"

        Write-Host "`r`n"

        if ($selection -eq '1')
        {
            $isQuit = ObjectSyncMenu

            if ($isQuit -eq $true)
            {
                break
            }
            else
            {
                continue
            }
        }
        elseif ($selection -eq '2')
        {
            $isQuit = PasswordSyncMenu

            if ($isQuit -eq $true)
            {
                break
            }
            else
            {
                continue
            }
        }
        elseif ($selection -eq '3')
        {
            Export-ADSyncDiagnosticsReport -OutputPath $ADSyncDiagnosticsOutputPath
        }
        elseif ($selection -eq '4')
        {
            if ($isNonInteractiveMode)
            {
                continue
            }

            if (-not (CheckConfigureAccountPermissionsPrerequisites))
            {
                continue
            }

            $isQuit = ConfigurePermissionsMenu

            if ($isQuit -eq $true)
            {
                break
            }
            else
            {
                continue
            }
        }
        elseif ($selection -eq '5')
        {
            if ($isNonInteractiveMode)
            {
                continue
            }

            Test-ADSyncAzureActiveDirectoryConnectivity
        }
        elseif ($selection -eq '6')
        {
            if ($isNonInteractiveMode)
            {
                continue
            }

            Test-ADSyncActiveDirectoryConnectivity
        }
        elseif ($selection -eq 'Q' -or $selection -eq 'q')
        {
            break
        }

        Write-Host "`r`n"
        Write-Host "`r`n"
    }
}

function Show-MainMenu
{
    Write-Host "`r`n"
    Write-Host "----------------------------------------AADConnect Troubleshooting------------------------------------------"
    Write-Host "`r`n"
    Write-Host "`tEnter '1' - Troubleshoot Object Synchronization"
    Write-Host "`tEnter '2' - Troubleshoot Password Hash Synchronization"
    Write-Host "`tEnter '3' - Collect General Diagnostics"
    Write-Host "`tEnter '4' - Configure AD DS Connector Account Permissions"
    Write-Host "`tEnter '5' - Test Azure Active Directory Connectivity"
    Write-Host "`tEnter '6' - Test Active Directory Connectivity"
    Write-Host "`tEnter 'Q' - Quit"
    Write-Host "`r`n"
}

#
# Options for Object Synchronization Troubleshooting
#
function ObjectSyncMenu
{
    $isQuit = $false
    
    while ($true)
    {
        Show-ObjectSyncMenu

        $selection = Read-Host "`tPlease make a selection"

        Write-Host "`r`n"

        if ($selection -eq '1')
        {
            Debug-ADSyncObjectSynchronizationIssues -DiagnosticOption $global:DiagnoseObjectSyncIssues
        }
        elseif ($selection -eq '2')
        {
            Debug-ADSyncAttributeSynchronizationIssues -DiagnosticOption $global:DiagnoseAttributeSyncIssues
        }
        elseif ($selection -eq '3')
        {
            Debug-ADSyncGroupMembershipSynchronizationIssues
        }
        elseif ($selection -eq '4')
        {
            Debug-ADSyncObjectSynchronizationIssues -DiagnosticOption $global:ChangePrimaryEmailAddress
        }
        elseif ($selection -eq '5')
        {
            Debug-ADSyncObjectSynchronizationIssues -DiagnosticOption $global:HideFromGlobalAddressList
        }
        elseif ($selection -eq '6')
        {
            Debug-ADSyncObjectAttributeRetrievalIssues
        }
        elseif ($selection -eq 'B' -or $selection -eq 'b')
        {
            Write-Output $isQuit
            break
        }
        elseif ($selection -eq 'Q' -or $selection -eq 'q')
        {
            $isQuit = $true
            Write-Output $isQuit
            break
        }

        Write-Host "`r`n"
        Write-Host "`r`n"
    }    
}

#
# Options for Password Hash Synchronization Troubleshooting
#
function PasswordSyncMenu
{
    $isQuit = $false

    while ($true)
    {
        Show-PasswordSyncMenu

        $selection = Read-Host "`tPlease make a selection"

        Write-Host "`r`n"

        if ($selection -eq '1')
        {
            DiagnosePasswordHashSync
        }
        elseif ($selection -eq 2)
        {
            DiagnosePasswordSyncSingleObject
        }
        elseif ($selection -eq 3)
        {
            SynchronizeSingleObjectPassword
        }
        elseif ($selection -eq 'B' -or $selection -eq 'b')
        {
            Write-Output $isQuit
            break
        }
        elseif ($selection -eq 'Q' -or $selection -eq 'q')
        {
            $isQuit = $true
            Write-Output $isQuit
            break
        }

        Write-Host "`r`n"
        Write-Host "`r`n"
    }
}

#
# Options for Configuring Permissions
#
function ConfigurePermissionsMenu
{
    $isQuit = $false
    
    while ($true)
    {
        Show-ConfigurePermissionsMenu

        $selection = Read-Host "`tPlease make a selection"

        Write-Host "`r`n"

        if ($selection -eq '1')
        {
            ConfigureAccountPermissions -ConfigurationOption "GetADSyncADConnectorAccount"
        }
        elseif ($selection -eq '2')
        {
            ConfigureAccountPermissions -ConfigurationOption "GetADSyncObjectsWithInheritanceDisabled"
        }
        elseif ($selection -eq '3')
        {
            ConfigureAccountPermissions -ConfigurationOption "SetADSyncBasicReadPermissions"
        }
        elseif ($selection -eq '4')
        {
            ConfigureAccountPermissions -ConfigurationOption "SetADSyncExchangeHybridPermissions"
        }
        elseif ($selection -eq '5')
        {
            ConfigureAccountPermissions -ConfigurationOption "SetADSyncExchangeMailPublicFolderPermissions"
        }
        elseif ($selection -eq '6')
        {
            ConfigureAccountPermissions -ConfigurationOption "SetADSyncMsDsConsistencyGuidPermissions"
        }
        elseif ($selection -eq '7')
        {
            ConfigureAccountPermissions -ConfigurationOption "SetADSyncPasswordHashSyncPermissions"
        }
        elseif ($selection -eq '8')
        {
            ConfigureAccountPermissions -ConfigurationOption "SetADSyncPasswordWritebackPermissions"
        }
        elseif ($selection -eq '9')
        {
            ConfigureAccountPermissions -ConfigurationOption "SetADSyncRestrictedPermissions"
        }
        elseif ($selection -eq '10')
        {
            ConfigureAccountPermissions -ConfigurationOption "SetADSyncUnifiedGroupWritebackPermissions"
        }
        elseif ($selection -eq '11')
        {
            ConfigureAccountPermissions -ConfigurationOption "ShowADSyncADObjectPermissions"
        }
        elseif ($selection -eq '12')
        {
            ConfigureAccountPermissions -ConfigurationOption "SetADSyncDefaultPermssions"
        }
        elseif ($selection -eq '13')
        {
            Debug-ADSyncObjectAttributeRetrievalIssues
        }
        elseif ($selection -eq 'B' -or $selection -eq 'b')
        {
            Write-Output $isQuit
            break
        }
        elseif ($selection -eq 'Q' -or $selection -eq 'q')
        {
            $isQuit = $true
            Write-Output $isQuit
            break
        }

        Write-Host "`r`n"
        Write-Host "`r`n"
    }    
}

function Show-ObjectSyncMenu
{
    Write-Host "`r`n"
    Write-Host "------------------------------------Troubleshoot Object Synchronization------------------------------------"
    Write-Host "`r`n"
    Write-Host "`tEnter '1' - Diagnose Object Synchronization Issues"
    Write-Host "`tEnter '2' - Diagnose Attribute Synchronization Issues"
    Write-Host "`tEnter '3' - Diagnose Group Membership Synchronization Issues"
    Write-Host "`tEnter '4' - How to change Exchange Online primary email address"
    Write-Host "`tEnter '5' - How to hide mailbox from Exchange Online global address list"
    Write-Host "`tEnter '6' - Compare object read permissions when running in context of AD Connector account vs Admin account"
    Write-Host "`tEnter 'B' - Go back to main troubleshooting menu"
    Write-Host "`tEnter 'Q' - Quit"
    Write-Host "`r`n"
}

function Show-PasswordSyncMenu
{
    Write-Host "`r`n"
    Write-Host "--------------------------------Troubleshoot Password Hash Synchronization---------------------------------"
    Write-Host "`r`n"
    Write-Host "`tEnter '1' - Password Hash Synchronization does NOT work at all"
    Write-Host "`tEnter '2' - Password Hash Synchronization does NOT work for a specific user account"
    Write-Host "`tEnter '3' - Synchronize password hash for a specific user account"
    Write-Host "`tEnter 'B' - Go back to main troubleshooting menu"
    Write-Host "`tEnter 'Q' - Quit"
    Write-Host "`r`n"
}

function Show-ConfigurePermissionsMenu
{
    Write-Host "`r`n"
    Write-Host "--------------------------------------------Configure Permissions------------------------------------------"
    Write-Host "`r`n"
    Write-Host "`tEnter '1' - Get AD Connector account"
    Write-Host "`tEnter '2' - Get objects with inheritance disabled"
    Write-Host "`tEnter '3' - Set basic read permissions"
    Write-Host "`tEnter '4' - Set Exchange Hybrid permissions"
    Write-Host "`tEnter '5' - Set Exchange mail public folder permissions"
    Write-Host "`tEnter '6' - Set MS-DS-Consistency-Guid permissions"
    Write-Host "`tEnter '7' - Set password hash sync permissions"
    Write-Host "`tEnter '8' - Set password writeback permissions"
    Write-Host "`tEnter '9' - Set restricted permissions"
    Write-Host "`tEnter '10' - Set unified group writeback permissions"
    Write-Host "`tEnter '11' - Show AD object permissions"
    Write-Host "`tEnter '12' - Set default AD Connector account permissions"
    Write-Host "`tEnter '13' - Compare object read permissions when running in context of AD Connector account vs Admin account"
    Write-Host "`tEnter 'B' - Go back to main troubleshooting menu"
    Write-Host "`tEnter 'Q' - Quit"
    Write-Host "`r`n"
}

Export-ModuleMember -Function Invoke-ADSyncDiagnostics
Export-ModuleMember -Function Invoke-ADSyncSingleObjectSync

# SIG # Begin signature block
# MIIoOQYJKoZIhvcNAQcCoIIoKjCCKCYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAkSvwNRNjHZAVe
# vfdvZvyYHFdwNQdJidaVsVX8NWbbBaCCDYIwggYAMIID6KADAgECAhMzAAADXJXz
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
# SEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGg0wghoJAgEBMIGVMH4xCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jv
# c29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAANclfNIW0oEas8AAAAAA1ww
# DQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIDEOS4K0
# VMhlx8aakP3dho4dsFSYq9+14/nxDeO8EsMoMEIGCisGAQQBgjcCAQwxNDAyoBSA
# EgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20w
# DQYJKoZIhvcNAQEBBQAEggEA1v0m4/R6uRpR9TagGdq+WiiyAtBG1sKMlcN3jNvg
# 3aolXwlZzil5IB0+yKNZYCFcrSPW3ary0iDkxMGduXfT4EXik5nd6Wl+r9BuYnsi
# PBA6KhMopsKYg5ZcWUY21z/P2dk7F9jiAZCbo2moL/rxflWelujAhhRSOhicJ6Ow
# Y0O9NzxiJNhEA9DCUU44OW9xIAfnvBPEUA3Jx88o9DtSMgWrcabxan5DuKf5xv00
# /VVYkWTs5dW7tJxfiSItL42c3DJa6KJuK2YN8er36OKPSMKmMploLww9KGilRblm
# RHpB6Gp5ZZQC5FjGId4ITzRt1iqLo85Id12wsuXo72kqY6GCF5cwgheTBgorBgEE
# AYI3AwMBMYIXgzCCF38GCSqGSIb3DQEHAqCCF3AwghdsAgEDMQ8wDQYJYIZIAWUD
# BAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoD
# ATAxMA0GCWCGSAFlAwQCAQUABCBaPIVcVhsO4SxFB7IXbV086B3vaUjxPn5BkD5Z
# LDvO9AIGZQP3SfWdGBMyMDIzMTAwNDE5Mjg1Ni4yNjdaMASAAgH0oIHRpIHOMIHL
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxN
# aWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRT
# UyBFU046OEQwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFNlcnZpY2WgghHtMIIHIDCCBQigAwIBAgITMwAAAc1VByrnysGZHQABAAAB
# zTANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAe
# Fw0yMzA1MjUxOTEyMDVaFw0yNDAyMDExOTEyMDVaMIHLMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmlj
# YSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OEQwMC0wNUUw
# LUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDTOCLVS2jmEWOqxzygW7s6YLmm
# 29pjvA+Ch6VL7HlTL8yUt3Z0KIzTa2O/Hvr/aJza1qEVklq7NPiOrpBAIz657LVx
# wEc4BxJiv6B68a8DQiF6WAFFNaK3WHi7TfxRnqLohgNz7vZPylZQX795r8MQvX56
# uwjj/R4hXnR7Na4Llu4mWsml/wp6VJqCuxZnu9jX4qaUxngcrfFT7+zvlXClwLah
# 2n0eGKna1dOjOgyK00jYq5vtzr5NZ+qVxqaw9DmEsj9vfqYkfQZry2JO5wmgXX79
# Ox7PLMUfqT4+8w5JkdSMoX32b1D6cDKWRUv5qjiYh4o/a9ehE/KAkUWlSPbbDR/a
# GnPJLAGPy2qA97YCBeeIJjRKURgdPlhE5O46kOju8nYJnIvxbuC2Qp2jxwc6rD9M
# 6Pvc8sZIcQ10YKZVYKs94YPSlkhwXwttbRY+jZnQiDm2ZFjH8SPe1I6ERcfeYX1z
# CYjEzdwWcm+fFZmlJA9HQW7ZJAmOECONtfK28EREEE5yzq+T3QMVPhiEfEhgcYsh
# 0DeoWiYGsDiKEuS+FElMMyT456+U2ZRa2hbRQ97QcbvaAd6OVQLp3TQqNEu0es5Z
# q0wg2CADf+QKQR/Y6+fGgk9qJNJW3Mu771KthuPlNfKss0B1zh0xa1yN4qC3zoE9
# Uq6T8r7G3/OtSFms4wIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFKGT+aY2aZrBAJVI
# Zh5kicokfNWaMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1Ud
# HwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3Js
# L01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggr
# BgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIw
# MTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgw
# DgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBSqG3ppKIU+i/EMwwt
# otoxnKfw0SX/3T16EPbjwsAImWOZ5nLAbatopl8zFY841gb5eiL1j81h4DiEiXt+
# BJgHIA2LIhKhSscd79oMbr631DiEqf9X5LZR3V3KIYstU3K7f5Dk7tbobuHu+6fY
# M/gOx44sgRU7YQ+YTYHvv8k4mMnuiahJRlU/F2vavcHU5uhXi078K4nSRAPnWyX7
# gVi6iVMBBUF4823oPFznEcHup7VNGRtGe1xvnlMd1CuyxctM8d/oqyTsxwlJAM5F
# /lDxnEWoSzAkad1nWvkaAeMV7+39IpXhuf9G3xbffKiyBnj3cQeiA4SxSwCdnx00
# RBlXS6r9tGDa/o9RS01FOABzKkP5CBDpm4wpKdIU74KtBH2sE5QYYn7liYWZr2f/
# U+ghTmdOEOPkXEcX81H4dRJU28Tj/gUZdwL81xah8Kn+cB7vM/Hs3/J8tF13ZPP+
# 8NtX3vu4NrchHDJYgjOi+1JuSf+4jpF/pEEPXp9AusizmSmkBK4iVT7NwVtRnS1t
# s8qAGHGPg2HPa4b2u9meueUoqNVtMhbumI1y+d9ZkThNXBXz2aItT2C99DM3T3qY
# qAUmvKUryVSpMLVpse4je5WN6VVlCDFKWFRH202YxEVWsZ5baN9CaqCbCS0Ea7s9
# OFLaEM5fNn9m5s69lD/ekcW2qTCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkA
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
# VTNYs6FwZvKhggNQMIICOAIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2Eg
# T3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjhEMDAtMDVFMC1E
# OTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEw
# BwYFKw4DAhoDFQBoqfem2KKzuRZjISYifGolVOdyBKCBgzCBgKR+MHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6MgpjzAiGA8y
# MDIzMTAwNDE4MTIzMVoYDzIwMjMxMDA1MTgxMjMxWjB3MD0GCisGAQQBhFkKBAEx
# LzAtMAoCBQDoyCmPAgEAMAoCAQACAgIlAgH/MAcCAQACAhNrMAoCBQDoyXsPAgEA
# MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAI
# AgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAKwmvXFQ1GHR/3e/7JOYI9dHfgTY
# 2RltTB9d4SOgVwKPqoJcK+Hii1/IPIqQ52M4vYJW70nd3RlO/E/Lk7yCDDqVBSrA
# ATFN036jhsKihn+mQSuN6UUVCKStiWRmElzQ9TefrHYi0ZfgfRU0Cd33U+nOHFWc
# OfkEWpzWCI9rw/HH9l6dOUPF2H4/26kcC51/kDvoQ+X9rtjk2vzuhylxENf0PJrv
# qlB4JXnqXzP/1UCHg47sZmJzhWOM9h/KRC8G5g8ZFH/2ei+ZGlYqBLWiIGqUyiGO
# thFGZsMP5EVeDXT3cx7LZkwqkFBAHa5GNRCo5iAWOzMkLZVI64KMptjLAnoxggQN
# MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAc1V
# ByrnysGZHQABAAABzTANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
# CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCDvO7+lALzvsx3feUi0GjV7AVCG
# pUVA4dfXZnuBZxhFljCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIOJmpfit
# Vr1PZGgvTEdTpStUc6GNh7LNroQBKwpURpkKMIGYMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAHNVQcq58rBmR0AAQAAAc0wIgQgWY7BoZ6J
# iDvx6d7Y6x35z7T/A50F6NtEUqU3u0crZOMwDQYJKoZIhvcNAQELBQAEggIAMziB
# TN7IGxlPhNvl422yuid+msGlIHs/7ujrWPBe/P6Fr1icwvmwch4XeP7jhFvk+5XX
# 9t9oauj36BdZKdIXFN5Tr1CBkHNm3V27gTSpr4TnduiqzIZ9oceJoxOL2JskltLD
# 5Ndbzupe5i9d2IkD5F6Gbt7UB0rvfNs/oRA2zLbjiZihGIwtHteNiY/QwOt3Q/+K
# j2hHfSOvIwJaaeDioZXJDYx8AyzfuU3QslUxld0+jXBQ6VfQ5najrdyIq8n+Y1Zw
# 3R5E8hFmW3TnvXdeV1/Kco0WdVt5lxdoXuw9LcWnG+X7DJz3uaWTydytiMCL22PO
# ZpxYN3wCnDzFwJP6K4b87n1uIwfA+WrKECy/zyRrkrTtI6sXSJKCnFh/gauARd/Q
# qdBlLHf6YnHhsR6Isjz9uH522bKGVjus8Gw59BaHWQAjJKsk3RHJv90XGupf0bTk
# vkXOkl3zGjsvDlxd3g71fg8PhNWeiL+1UGchPikM8dU/NyhmqK8iCUnoC0QQIToc
# ISSTBlt2T3ICwzN1QQRoY49MzqsN4yoA6uD1sxmbapTF9AFAaA+OQ/PCEvD0rlaT
# mzPy3af216T7CNnTUerWjOX2OOkvn1h7OLTWrkTXjxTkCrAGfefPgUNR229Iz0t6
# QpgWyWm9uVj/m6v7GFuozO6+vhY9lXfnbf5MB30=
# SIG # End signature block
