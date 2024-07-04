#-------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation.  All rights reserved.
#-------------------------------------------------------------------------

Function CheckConfigureAccountPermissionsPrerequisites
{
    # Check if all the features installed by RSAT-AD-Tools are available. These are required for the AdSyncConfig module
    if (((Get-WindowsFeature RSAT-AD-PowerShell).InstallState -ne "Installed") -or ((Get-WindowsFeature RSAT-AD-AdminCenter).InstallState -ne "Installed") -or ((Get-WindowsFeature RSAT-ADDS-Tools).InstallState -ne "Installed") -or ((Get-WindowsFeature RSAT-ADLDS).InstallState -ne "Installed"))
    {
        $installRSATOptions = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes", "&No")
        $canInstall = !($host.UI.PromptForChoice("RSAT-AD-Tools Warning", "The use of these options requires 'Remote Server Administration Tools for AD DS' to be installed. Is it okay to install them now?", $installRSATOptions, 0))

        if (!$canInstall)
        {
            return $false
        }

        Try
        {
            $installState = $null

            $osVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductName).ProductName

            if ($osVersion -Like "*Windows Server 2008 R2*")
            {
                $installState = Add-WindowsFeature RSAT-AD-Tools
            }
            else
            {
                $installState = Install-WindowsFeature RSAT-AD-Tools
            }

            if ($installState.RestartNeeded -eq "Yes")
            {
                $restartOptions = [System.Management.Automation.Host.ChoiceDescription[]] @("&Restart Now", "&Back to Previous Menu")
                $confirmRestart = !($host.UI.PromptForChoice("Feature Install Warning", "A restart is required to complete installation of the RSAT-AD-Tools Windows Feature. Is it okay to restart now?", $installRSATOptions, 0))

                if ($confirmRestart)
                {
                    Restart-Computer
                }
                else
                {
                    return $false
                }
            }
        }
        Catch
        {
            Write-Host "[ERROR]`t Installing RSAT-AD-Tools Windows Feature failed."
            Exit 1
        }
    }

    if (-not (Get-Module ActiveDirectory))
    {
        Import-Module ActiveDirectory -ErrorAction Stop
    }

    # These options make use of the AdSyncConfig module included with the product
    Try
    {
        $aadConnectRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Azure AD Connect"
        $aadConnectWizardPath = $aadConnectRegKey.WizardPath
        $aadConnectWizardFileName = "AzureADConnect.exe"

        $aadConnectPathLength = $aadConnectWizardPath.IndexOf($aadConnectWizardFileName)
        $aadConnectPath = $aadConnectWizardPath.Substring(0, $aadConnectPathLength)

        $adSyncConfigModulePath = [System.IO.Path]::Combine($aadConnectPath, "AdSyncConfig\AdSyncConfig.psm1")
        Import-Module $adSyncConfigModulePath -ErrorAction Stop
    }
    Catch
    {
        Write-Host "[ERROR]`t Importing the AdSyncTools module failed."
        Exit 1
    }

    return $true
}

Function ConfigureAccountPermissions
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $ConfigurationOption
    )

    Write-Host "`r`n"

    if ($ConfigurationOption -eq "GetADSyncADConnectorAccount")
    {
        $allAccounts = Get-ADSyncADConnectorAccount

        Write-Host ($allAccounts | Format-Table -Wrap -AutoSize -Property ADConnectorName, ADConnectorForest, ADConnectorAccountName,ADConnectorAccountDomain | Out-String)
    }
    elseif ($ConfigurationOption -eq "GetADSyncObjectsWithInheritanceDisabled")
    {
        Write-Host "`r`n"
        $targetForest = Read-Host "Please enter the forest to search for objects"
        Write-Host "`r`n"
        $searchBase = Read-Host "Please enter the search base for the LDAP query, it can be an AD Domain DistinguishedName or a FQDN"
        Write-Host "`r`n"
        $objectClass = Read-Host "Please enter the object class. It can be '*' (for any object class), 'user', 'group', 'container', etc. Giving no input will use the default, 'organizationalUnit'"

        if ($targetForest -ne (Get-ADForest).Name)
        {
            $targetForestCredentials = Get-Credential -Message "Please enter credentials for target forest '$targetForest' in DOMAIN\username format"

            if ($objectClass)
            {
                Write-Host (Get-ADSyncObjectsWithInheritanceDisabled -SearchBase $searchBase -ObjectClass $objectClass -TargetForest $targetForest -TargetForestCredential $targetForestCredentials | Out-String)
            }
            else
            {
                Write-Host (Get-ADSyncObjectsWithInheritanceDisabled -SearchBase $searchBase -TargetForest $targetForest -TargetForestCredential $targetForestCredentials | Out-String)
            }
        }
        else
        {
            if ($objectClass)
            {
                Write-Host (Get-ADSyncObjectsWithInheritanceDisabled -SearchBase $searchBase -ObjectClass $objectClass | Out-String)
            }
            else
            {
                Write-Host (Get-ADSyncObjectsWithInheritanceDisabled -SearchBase $searchBase | Out-String)
            }
        }
    }
    elseif ($ConfigurationOption -eq "ShowADSyncADObjectPermissions")
    {
        Write-Host "`r`n"
        $targetForest = Read-Host "Please enter the forest where the object exists"

        $objectToCheck = Read-Host "Please enter the DistinguishedName of the object to view the permissions on"

        if ($targetForest -ne (Get-ADForest).Name)
        {
            $targetForestCredentials = Get-Credential -Message "Please enter credentials for target forest '$targetForest' in DOMAIN\username format"

            Write-Host (Show-ADSyncADObjectPermissions -ADobjectDN $objectToCheck -TargetForestCredential $targetForestCredentials | Out-String)
        }
        else
        {
            Write-Host (Show-ADSyncADObjectPermissions -ADobjectDN $objectToCheck | Out-String)
        }
    }
    else
    {
        if ($ConfigurationOption -eq "SetADSyncDefaultPermssions")
        {
            Write-Host "`r`n"
            Write-Host "This option will set permissions required for the following:"
            Write-Host "    Password Hash Sync"
            Write-Host "    Password Writeback"
            Write-Host "    Hybrid Exchange"
            Write-Host "    Exchange Mail Public Folder"
            Write-Host "    MsDsConsistencyGuid"
            Write-Host "It will then restrict permissions"

            $confirmDefaultOptions = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes", "&No")
            $confirmDefaultChoice = !($host.UI.PromptForChoice("Confirm", "Would you like to continue with these options?", $confirmDefaultOptions, 0))

            if (!$confirmDefaultChoice)
            {
                return
            }
        }

        $configureArguments = @{ }

        $accountTypeOptions = [System.Management.Automation.Host.ChoiceDescription[]] @("&Existing Connector Account", "&Custom Account")
        $accountTypeChoice = $host.UI.PromptForChoice("Account to Configure", "Would you like to configure an existing connector account or a custom account?", $accountTypeOptions, 0)

        if ($accountTypeChoice)
        {
            Write-Host "`r`n"
            $accountChoiceOptions = [System.Management.Automation.Host.ChoiceDescription[]] @("&Distinguished Name", "&Account Name and Account Domain")
            $accountInputChoice = $host.UI.PromptForChoice("Account Information", "How would you like to input the account information?", $accountChoiceOptions, 0)

            Write-Host "`r`n"
            $targetForest = Read-Host "Please enter the forest on which to configure the account"

            if ($accountInputChoice)
            {
                $connectorAccountName = Read-Host "Please enter the Account Name"
                $connectorAccountDomain = Read-Host "Please enter the Account Domain"

                $configureArguments.Add("ADConnectorAccountName", $connectorAccountName)
                $configureArguments.Add("ADConnectorAccountDomain", $connectorAccountDomain)
            }
            else
            {
                Write-Host "`r`n"
                $connectorAccountDN = Read-Host "Please enter the Account Distinguished Name"

                $configureArguments.Add("ADConnectorAccountDN", $connectorAccountDN)
            }
        }
        else
        {
            $allAccounts = Get-ADSyncADConnectorAccount

            $selectedConnector = $null

            while ($selectedConnector -eq $null)
            {
                Write-Host "`r`n"
                Write-Host "Configured connectors and their related accounts:"

                Write-Host ($allAccounts | Out-String)

                Write-Host "`r`n"
                $connectorSelection = Read-Host "Name of the connector who's account to configure"

                if ($allAccounts.ADConnectorName -contains $connectorSelection)
                {
                    $selectedConnector = $allAccounts | Where-Object { $_.ADConnectorName -like $connectorSelection }
                    $configureArguments.Add("ADConnectorAccountName", $selectedConnector.ADConnectorAccountName)
                    $configureArguments.Add("ADConnectorAccountDomain", $selectedConnector.ADConnectorAccountDomain)

                    $targetForest = $selectedConnector.ADConnectorForest
                }
                else
                {
                    Write-Host "`r`n"
                    ReportWarning "'$connectorSelection' is not a configured connector. Please try again!"
                }
            }
        }

        if ($targetForest -ne (Get-ADForest).Name)
        {
            $targetForestCredentials = Get-Credential -Message "Please enter credentials for target forest '$targetForest' in DOMAIN\username format"

            $configureArguments.Add("TargetForest", $targetForest)
            $configureArguments.Add("TargetForestCredential", $targetForestCredentials)
        }

        Write-Host "`r`n"

        if (($ConfigurationOption -eq "SetADSyncRestrictedPermissions") -or ($ConfigurationOption -eq "SetADSyncDefaultPermssions"))
        {
            if ($configureArguments.ContainsKey("ADConnectorAccountName"))
            {
                $ldapFilter = "(|(Name=" +  $configureArguments.ADConnectorAccountName + ")(sAMAccountName=" + $configureArguments.ADConnectorAccountName + "))"

                if ($targetForest -ne (Get-ADForest).Name)
                {
                    $adObject = Get-ADObject -LDAPFilter $ldapFilter -Server $configureArguments.ADConnectorAccountDomain -Properties distinguishedName -Credential $configureArguments.TargetForestCredential -ErrorAction Stop
                }
                else
                {
                    $adObject = Get-ADObject -LDAPFilter $ldapFilter -Server $configureArguments.ADConnectorAccountDomain -Properties distinguishedName -ErrorAction Stop
                }

                $configureArguments.Remove("ADConnectorAccountName")
                $configureArguments.Remove("ADConnectorAccountDomain")
                $configureArguments.Add("ADConnectorAccountDN", $adObject.distinguishedName)
            }

            $adminCredential = Get-Credential -Message "Please enter Administrator credentials that have the necessary privileges to restrict the permissions on the Connector Account"

            if ($ConfigurationOption -eq "SetADSyncDefaultPermssions")
            {
                Write-Host (Set-ADSyncPasswordHashSyncPermissions @configureArguments | Out-String)
                Write-Host (Set-ADSyncPasswordWritebackPermissions @configureArguments | Out-String)
                Write-Host (Set-ADSyncExchangeHybridPermissions @configureArguments | Out-String)
                Write-Host (Set-ADSyncExchangeMailPublicFolderPermissions @configureArguments | Out-String)
                Write-Host (Set-ADSyncMsDsConsistencyGuidPermissions @configureArguments | Out-String)

                $configureArguments.Add("Credential", $adminCredential)
                if ($configureArguments.ContainsKey("TargetForestCredential"))
                {
                    $configureArguments.Remove("TargetForestCredential")
                }

                Write-Host (Set-ADSyncRestrictedPermissions @configureArguments | Out-String)
            }
            else
            {
                $configureArguments.Add("Credential", $adminCredential)
                if ($configureArguments.ContainsKey("TargetForestCredential"))
                {
                    $configureArguments.Remove("TargetForestCredential")
                }

                Write-Host "`r`n"
                $validationChoiceOptions = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes", "&No")
                $disableCredentialValidation = !($host.UI.PromptForChoice("Disable Credential Validation", "Skip checking if the credentials provided are valid in AD and if they have the necessary privileges to restrict the permissions on the Connector account.", $validationChoiceOptions, 1))
        
                if ($disableCredentialValidation)
                {
                    $configureArguments.Add("DisableCredentialValidation", $true)
                }

                Write-Host "`r`n"
                Write-Host (Set-ADSyncRestrictedPermissions @configureArguments | Out-String)
            }
        }
        elseif ($ConfigurationOption -eq "SetADSyncPasswordHashSyncPermissions")
        {
            Write-Host (Set-ADSyncPasswordHashSyncPermissions @configureArguments | Out-String)
        }
        else
        {
            $adObjectDN = Read-Host "To set permissions for a single target object, enter the DistinguishedName of the target AD object. Giving no input will set root permissions for all Domains in the Forest"

            Write-Host "`r`n"
            $includeChoiceOptions = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes", "&No")
            $includeAdminSDHolders = !($host.UI.PromptForChoice("Update AdminSdHolders", "Update AdminSDHolder container when updating with these permissions?", $includeChoiceOptions, 1))

            if ($adObjectDN)
            {
                $configureArguments.Add("ADobjectDN", $adObjectDN)
            }

            if ($includeAdminSDHolders)
            {
                $configureArguments.Add("IncludeAdminSDHolders", $true)
            }

            Write-Host "`r`n"

            if ($ConfigurationOption -eq "SetADSyncBasicReadPermissions")
            {
                Write-Host (Set-ADSyncBasicReadPermissions @configureArguments | Out-String)
            }
            elseif ($ConfigurationOption -eq "SetADSyncExchangeHybridPermissions")
            {
                Write-Host (Set-ADSyncExchangeHybridPermissions @configureArguments | Out-String)
            }
            elseif ($ConfigurationOption -eq "SetADSyncExchangeMailPublicFolderPermissions")
            {
                Write-Host (Set-ADSyncExchangeMailPublicFolderPermissions @configureArguments | Out-String)
            }
            elseif ($ConfigurationOption -eq "SetADSyncMsDsConsistencyGuidPermissions")
            {
                Write-Host (Set-ADSyncMsDsConsistencyGuidPermissions @configureArguments | Out-String)
            }
            elseif ($ConfigurationOption -eq "SetADSyncPasswordWritebackPermissions")
            {
                Write-Host (Set-ADSyncPasswordWritebackPermissions @configureArguments | Out-String)
            }
            elseif ($ConfigurationOption -eq "SetADSyncUnifiedGroupWritebackPermissions")
            {
                Write-Host (Set-ADSyncUnifiedGroupWritebackPermissions @configureArguments | Out-String)
            }
        }
    }
}
# SIG # Begin signature block
# MIIoNgYJKoZIhvcNAQcCoIIoJzCCKCMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDZ7I+XwaUhLp33
# bPHQpAN5xYlDd9kR8siFMxxQw8Wbp6CCDYIwggYAMIID6KADAgECAhMzAAADXJXz
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJ2efEzX
# txOc3nQ2Ll5GC0G90qbA8lFkzofh667vytDUMEIGCisGAQQBgjcCAQwxNDAyoBSA
# EgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20w
# DQYJKoZIhvcNAQEBBQAEggEAX8l3bCyCy7MyrLzPHCRXS13+qRyeV7Zn+Hiyqllw
# I+7bTERyxI8lMNWGPsZNRmhSaj0P05qq7GbPc6wM+GUJnjWSLnxreaBX+UK3WVre
# yMxmNuXFbe7rEUBW6wrCYj3elBPlnYl8lFdMbdzQCtJ/M6XN3iScXlwgC+tkHSdG
# cuIRCZ/0/hDvnVV5Yq5pekW00bqBznPM/51v1jFH5VDvW7AvRsdMbgqaHkm5QWwQ
# wp8U2EV9bPOcNIGkt2FBJLXIsvi+mJ0sx1FbwASt6EDbKaCN1AHElj7gdfJvkQ/5
# QZX05rJoB+0ZyCfF9K3XRDnbXqrbN1UaY/0ASjyQPoM1SaGCF5QwgheQBgorBgEE
# AYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUD
# BAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoD
# ATAxMA0GCWCGSAFlAwQCAQUABCDsCgYHG90YX0JlevN7Z2ZGtGSw2HX4MzLm007I
# +bM9BgIGZQQ0jkCqGBMyMDIzMTAwNDE5MzA0OC40NDFaMASAAgH0oIHRpIHOMIHL
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
# SIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCBxNVWYaqiJwRHePBYWaUIKuXzX+Bv8
# 9g5ECBpB0mY7szCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIDJsz0F6L0XD
# Um53JRBfNMZszKsllLDMiaFZ3PL/LqxnMIGYMIGApH4wfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAHODxj3RZfnxv8AAQAAAc4wIgQg1+caYgNLKhNy
# gXcmWt0L0jo/QJsgWkeqnT9LPZdM47kwDQYJKoZIhvcNAQELBQAEggIAgcq5rVQA
# thNBFUaZ6tNcR6iEXfjw2ROOllzRL7HmxFusMFCBVCzwRfrGFU11jQr9pwicKBh9
# KqmgO05IzPnPuMg+x1TblwxtHv4X/bDZQ1g5SqiXq2DzFg8t1ga8aOTAiQyQRkAm
# J7eMVhEslyr69ln816Oi96B1MoZ7YHAPQDF6kTfhE22LEir1OCyGGy9coE5/LLlU
# kqq8gdkDoeE28Os2mfllvXJebCitiHO0pq5sIZwnP27ruiy7lwRugFq94U25+FkN
# 7zx297Nci+it8GZEXxxsofvZ+3Tkc34bPu+E5evdMu1313L56jeNbAwvzokkC5/i
# B7e1zE4pToGvjqoUhH4kVjluHeFYIjfCwuZD9mI0OjA+cjps1rk7WebI3mLSyTHU
# hZbSrhwWfaVm+07zIMW9jyEUnOC1hlzm9BWGL9WA9Sln600LG+liCou6ZUehpkgO
# cOvM2klNuWErZf+5yx5xCi9BEBSL3fd9xCnxQah6QkrsNgrcoN8QVQN0vYUBmILF
# kedj80rkHkB0bIHU7dzCyOEMddByiprlKnAcilBPji/LIrYkYMyKIvQWUTjqlzQK
# 8MygxvlmZ0Wo3t3I149LOUvYKyZH7cEnLJ41bSqv1bLqNFquGxYMrqXrd4Ee5ynG
# tOfv08OfMpDnFjMvNR7gGAZPhDops++pB4k=
# SIG # End signature block
