#-------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation.  All rights reserved.
#-------------------------------------------------------------------------

$ErrorActionPreference = "Stop"

$paramsRegKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters\'
if (Test-Path $paramsRegKey)
{
  $installationDirectory = (Get-ItemProperty $paramsRegKey -Name Path).Path
  Import-Module "$installationDirectory\Bin\ADSync\ADSync.psd1"
  Add-Type -Path "$installationDirectory\Bin\Assemblies\Microsoft.MetadirectoryServicesEx.dll"
  Add-Type -Path "$installationDirectory\Bin\Microsoft.MetadirectoryServices.PasswordHashSynchronization.Types.dll"
  Add-Type -Path "$installationDirectory\Extensions\Microsoft.Azure.ActiveDirectory.Connector.dll"
}
else 
{
  Write-Warning -Message "AADConnect installation was not found. Using current directory. Some functionality may be unavailable."
  Add-Type -Path "$PSScriptRoot\Microsoft.MetadirectoryServicesEx.dll"
  Add-Type -Path "$PSScriptRoot\Microsoft.MetadirectoryServices.PasswordHashSynchronization.Types.dll"
  Add-Type -Path "$PSScriptRoot\Microsoft.Azure.ActiveDirectory.Connector.dll"
}

function Get-ADSyncAADConnectorImportApiVersion
{
  $paramRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters"
  $paramRegName = "AadConnectorImportApiVersion"

  $version = (Get-ItemProperty -Path $paramRegKey | Select $paramRegName).$paramRegName
  if (-not $version)
  {
    $version = 2
  }

  Write-Output $version
}

function Get-ADSyncAADConnectorExportApiVersion
{
  $aad = Get-ADSyncConnector -Identifier b891884f-051e-4a83-95af-2544101c9083
  $aadRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters\PerMAInstance\" + $aad.Name
  $aadRegName = "EnableNonExistentObjectReferenceExportFiltering"
  $paramRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters"
  $paramRegName = "AadConnectorExportApiVersion"

  $exportType = $aad.ExtensionConfiguration.ExportType
  $enableNonExistentObjectReferenceExportFiltering = 1
  if (Test-Path -Path $aadRegKey)
  {
    $enableNonExistentObjectReferenceExportFiltering = (Get-ItemProperty -Path $aadRegKey | Select $aadRegName).$aadRegName
    if ([string]::IsNullOrEmpty($enableNonExistentObjectReferenceExportFiltering))
    {
      $enableNonExistentObjectReferenceExportFiltering = 1
    }
  }
  $version = (Get-ItemProperty -Path $paramRegKey | Select $paramRegName).$paramRegName
  if (-not $version)
  {
    $version = 2
  }

  if ($version -eq 1)
  {
    if ($exportType -ne 2 -or $enableNonExistentObjectReferenceExportFiltering -ne 0)
    {
      Write-Error "Version mismatch in Registry and AAD Connector. Run Set-ADSyncAADConnectorExportApiVersion to fix the issue"
      return
    }
  }
  elseif ($version -eq 2)
  {
    if ($exportType -ne 5 -or $enableNonExistentObjectReferenceExportFiltering -ne 1)
    {
      Write-Error "Version mismatch in Registry and AAD Connector. Run Set-ADSyncAADConnectorExportApiVersion to fix the issue"
      return
    }
  }
  else
  {
    Write-Error "unknown version $version"
    return
  }

  Write-Output $version
}

function Set-ADSyncAADConnectorImportApiVersion 
{
  Param
  (
    [Parameter(Mandatory=$true)]
    [int] $version
  )

  $regKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters"
  $regName = "AadConnectorImportApiVersion"

  if (($version -ne 1) -and ($version -ne 2))
  {
    Write-Error "unknown version $version"
  }

  Set-ItemProperty -path $regKey -name $regName -value $version

  Write-Output $version
}

function Set-ADSyncAADConnectorExportApiVersion
{
  Param
  (
    [Parameter(Mandatory=$true)]
    [int] $version
  )

  $paramRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters"
  $regName = "AadConnectorExportApiVersion"
  $exportType = $null;
  $enableNonExistentObjectReferenceExportFiltering = 1;

  if ($version -eq 1)
  {
    $exportType = 2 
    $enableNonExistentObjectReferenceExportFiltering = 0
  }
  elseif ($version -eq 2)
  {   
    $exportType = 5
    $enableNonExistentObjectReferenceExportFiltering = 1
  }
  else
  {
    Write-Error "unknown version $version"
  }

  Set-ItemProperty -path $paramRegKey -name $regName -value $version
  $aad = Get-ADSyncConnector -Identifier b891884f-051e-4a83-95af-2544101c9083

  $perMaInstanceRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters\PerMAInstance\"
  if (-not (Test-Path -Path $perMaInstanceRegKey))
  {
    New-Item -Path $perMaInstanceRegKey |  Out-Null
  }
  $aadRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters\PerMAInstance\" + $aad.Name
  if (-not (Test-Path -Path $aadRegKey))
  {
    New-Item -Path $aadRegKey |  Out-Null
  }
  Set-ItemProperty -Path $aadRegKey -Name "EnableNonExistentObjectReferenceExportFiltering" -Value $enableNonExistentObjectReferenceExportFiltering

  $aad.ExtensionConfiguration.ExportType = $exportType
  $aad = Add-ADSyncConnector -Connector $aad

  Write-Output $version
}

function ConvertTo-ADSyncSourceAnchor
{
  Param
  (
    [Parameter(Mandatory=$true)]
    [string] $dn
  )

  $result = [Microsoft.Online.DirSync.Extension.Utilities.DNEncoding]::SafeRdnToString($dn);
  Write-Output $result
}

function ConvertTo-ADSyncAadDistinguishedName
{
  Param
  (
    [Parameter(Mandatory=$true)]
    [string] $sourceAnchor
  )

  $result = [Microsoft.Online.DirSync.Extension.Utilities.DNEncoding]::StringToSafeRdn($sourceAnchor);
  Write-Output $result
}

function ConvertTo-ADSyncCloudAnchor
{
  Param
  (
    [Parameter(Mandatory=$true)]
    [string] $encodedB64Anchor
  )
  $encodedRawAnchor =  [System.Convert]::FromBase64String($encodedB64Anchor);
  $rawAnchor = $encodedRawAnchor[4..($encodedRawAnchor.Length - 3)]
  $cloudAnchor = [System.Text.Encoding]::Unicode.GetString($rawAnchor)
  Write-Output $cloudAnchor
}

function Set-ADSyncMaxImportRefetchSize
{
  Param
  (
    [Parameter(Mandatory=$true)]
    [int] $size
  )

  $paramRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters"
  $regName = "MaxImportRefetchSize"

  Set-ItemProperty -path $paramRegKey -name $regName -value $size

  Write-Output $size
}

function Enable-ADSyncAADConnectorReferenceObsoletion
{
  $aadConnector = Get-ADSyncConnector -Identifier b891884f-051e-4a83-95af-2544101c9083

  $perMaInstanceRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters\PerMAInstance\"
  if (-not (Test-Path -Path $perMaInstanceRegKey))
  {
    New-Item -Path $perMaInstanceRegKey |  Out-Null
  }
  $aadConnectorRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters\PerMAInstance\" + $aadConnector.Name
  if (-not (Test-Path -Path $aadConnectorRegKey))
  {
    New-Item -Path $aadConnectorRegKey |  Out-Null
  }
  Set-ItemProperty -Path $aadConnectorRegKey -Name "EnableReferenceObsoletion" -Value 1

  Write-Output $true
}

function Disable-ADSyncAADConnectorReferenceObsoletion
{
  $aadConnector = Get-ADSyncConnector -Identifier b891884f-051e-4a83-95af-2544101c9083

  $perMaInstanceRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters\PerMAInstance\"
  if (-not (Test-Path -Path $perMaInstanceRegKey))
  {
    New-Item -Path $perMaInstanceRegKey |  Out-Null
  }
  $aadConnectorRegKey = "HKLM:\SYSTEM\CurrentControlSet\Services\ADSync\Parameters\PerMAInstance\" + $aadConnector.Name
  if (-not (Test-Path -Path $aadConnectorRegKey))
  {
    New-Item -Path $aadConnectorRegKey |  Out-Null
  }
  Set-ItemProperty -Path $aadConnectorRegKey -Name "EnableReferenceObsoletion" -Value 0

  Write-Output $true
}

#
# Sample usage:
#   Full import: Get-ADSyncAadDirectoryChanges -creds $creds -version 2
#
function Get-ADSyncAadDirectoryChanges
{
  Param
  (
    [Parameter(Mandatory=$true)]
    [PSCredential] $creds,

    [Parameter(Mandatory=$false)]
    [byte[]] $watermark,

    [Parameter(Mandatory=$false)]
    [int] $version = 1,

    [Parameter(Mandatory=$false)]
    [string[]] $objectClassFilter
  )

  $enumerator = `
    [Microsoft.Azure.ActiveDirectory.Connector.Diagnostics.DiagnosticsFactory]::CreateChangeEnumerator( `
        $creds.UserName, `
        $creds.Password, 
        $objectClassFilter, `
        $watermark, `
        ($watermark -eq $null), `
        $version)

  try {
      while ($true) {
        $res = $enumerator.EnumerateNextBatch()
        Write-Output $res
        if (-not $res.AadBatch.MoreToRead) {
            break;
        }
    }
  }
  finally {
    $enumerator.Dispose()
  }
}

#
# Sample usage: 
# Collect AAD user object DNs to delete in file userDNs.txt and run:
#   Remove-ADSyncAadObjects -creds $creds -entries $( type .\userDNs.txt | ForEach-Object { [Microsoft.Azure.ActiveDirectory.Connector.Diagnostics.DeletionEntry]::FromDistinguishedName("user", $_) } )
#
function Remove-ADSyncAadObjects
{
  Param
  (
    [Parameter(Mandatory=$true)]
    [PSCredential] $creds,

    [Parameter(Mandatory=$true)]
    [Microsoft.Azure.ActiveDirectory.Connector.Diagnostics.DeletionEntry[]] $entries
  )

  $len = $entries.Count
  $entriesProcessed = 0
  $batchSize = 10

  $exporter = `
    [Microsoft.Azure.ActiveDirectory.Connector.Diagnostics.DiagnosticsFactory]::CreateDirectoryChangeExporter( `
      $creds.UserName, `
      $creds.Password)

  try {
    while ($entriesProcessed -lt $len) {
      $nextBatch = $entries | Select-Object -First $batchSize
      $entries = $entries | Select-Object -Skip $batchSize

      $results = $exporter.ExportDeletions([Microsoft.Azure.ActiveDirectory.Connector.Diagnostics.DeletionEntry[]]$nextBatch)
      Write-Output $results
  
      $entriesProcessed += $nextBatch.Count
    }
  } finally {
    $exporter.Dispose()
  }
}

# SIG # Begin signature block
# MIIoNgYJKoZIhvcNAQcCoIIoJzCCKCMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBnSNOH//dxFJEH
# 2xhnTU7oLyyJAofe9EQ2s8CBSTa1HKCCDYIwggYAMIID6KADAgECAhMzAAADXJXz
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIH7etgK0
# 1m0cpKZLIFLzKuZUbQG5dQtqPRa/kVGcNYnpMEIGCisGAQQBgjcCAQwxNDAyoBSA
# EgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20w
# DQYJKoZIhvcNAQEBBQAEggEAevUJ+qtx0CQ9tZtASDEk4WfDBv71ycLk7WpFg+H5
# I244PTHu6stIIqNCl6/KSSyi2IhhcxQQ0A/b2vhoL+lhMLuijo8vH65q+E6/mT0h
# cz0dDzS9LjSlM2j6YZWqeLZTP7K+G1kncdb+Bj0LWnn4enkAUFpCUs9rB8ss2Ea1
# 9u3uaZ+1XaGCJEQkwLuo1ga6LCeqLJn63EXLjcwW5oSTwgXdkcnT2ieIsgAgHgsc
# 9tpsFRMNbfAtAonL4OAEf9AMgMfyUxRB3dn3eHbjG9xDhg/AKAhhKGFSIw10C5pd
# oX7EIrIuRQ45jxHETXwZER0laYk/Ubm+c19nCIPTSWJI5aGCF5QwgheQBgorBgEE
# AYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUD
# BAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoD
# ATAxMA0GCWCGSAFlAwQCAQUABCCtSBwBaQ5nRcWyHyUzXj3VOKPCuDVTVnsh3qZk
# aWm35wIGZQtlabC2GBMyMDIzMTAwNDE5MjgyMi43MzFaMASAAgH0oIHRpIHOMIHL
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxN
# aWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRT
# UyBFU046MzcwMy0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAdTk6QMvwKxprAABAAAB
# 1DANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAe
# Fw0yMzA1MjUxOTEyMjdaFw0yNDAyMDExOTEyMjdaMIHLMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmlj
# YSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046MzcwMy0wNUUw
# LUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCYU94tmwIkl353SWej1ybWcSAb
# u8FLwTEtOvw3uXMpa1DnDXDwbtkLc+oT8BNti8t+38TwktfgoAM9N/BOHyT4CpXB
# 1Hwn1YYovuYujoQV9kmyU6D6QttTIKN7fZTjoNtIhI5CBkwS+MkwCwdaNyySvjwP
# vZuxH8RNcOOB8ABDhJH+vw/jev+G20HE0Gwad323x4uA4tLkE0e9yaD7x/s1F3lt
# 7Ni47pJMGMLqZQCK7UCUeWauWF9wZINQ459tSPIe/xK6ttLyYHzd3DeRRLxQP/7c
# 7oPJPDFgpbGB2HRJaE0puRRDoiDP7JJxYr+TBExhI2ulZWbgL4CfWawwb1LsJmFW
# JHbqGr6o0irW7IqDkf2qEbMRT1WUM15F5oBc5Lg18lb3sUW7kRPvKwmfaRBkrmil
# 0H/tv3HYyE6A490ZFEcPk6dzYAKfCe3vKpRVE4dPoDKVnCLUTLkq1f/pnuD/ZGHJ
# 2cbuIer9umQYu/Fz1DBreC8CRs3zJm48HIS3rbeLUYu/C93jVIJOlrKAv/qmYRym
# jDmpfzZvfvGBGUbOpx+4ofwqBTLuhAfO7FZz338NtsjDzq3siR0cP74p9UuNX1Tp
# z4KZLM8GlzZLje3aHfD3mulrPIMipnVqBkkY12a2slsbIlje3uq8BSrj725/wHCt
# 4HyXW4WgTGPizyExTQIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFDzajMdwtAZ6EoB5
# Hedcsru0DHZJMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1Ud
# HwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3Js
# L01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggr
# BgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIw
# MTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgw
# DgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQC0xUPP+ytwktdRhYlZ
# 9Bk4/bLzLOzq+wcC7VAaRQHGRS+IPyU/8OLiVoXcoyKKKiRQ7K9c90OdM+qL4Piz
# KnStLDBsWT+ds1hayNkTwnhVcZeA1EGKlNZvdlTsCUxJ5C7yoZQmA+2lpk04PGjc
# FhH1gGRphz+tcDNK/CtKJ+PrEuNj7sgmBop/JFQcYymiP/vr+dudrKQeStcTV9W1
# 3cm2FD5F/XWO37Ti+G4Tg1BkU25RA+t8RCWy/IHug3rrYzqUcdVRq7UgRl40YIkT
# Nnuco6ny7vEBmWFjcr7Skvo/QWueO8NAvP2ZKf3QMfidmH1xvxx9h9wVU6rvEQ/P
# UJi3popYsrQKuogphdPqHZ5j9OoQ+EjACUfgJlHnn8GVbPW3xGplCkXbyEHheQNd
# /a3X/2zpSwEROOcy1YaeQquflGilAf0y40AFKqW2Q1yTb19cRXBpRzbZVO+RXUB4
# A6UL1E1Xjtzr/b9qz9U4UNV8wy8Yv/07bp3hAFfxB4mn0c+PO+YFv2YsVvYATVI2
# lwL9QDSEt8F0RW6LekxPfvbkmVSRwP6pf5AUfkqooKa6pfqTCndpGT71Hyiltela
# MhRUsNVkaKzAJrUoESSj7sTP1ZGiS9JgI+p3AO5fnMht3mLHMg68GszSH4Wy3vUD
# JpjUTYLtaTWkQtz6UqZPN7WXhjCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkA
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
# T3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjM3MDMtMDVFMC1E
# OTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEw
# BwYFKw4DAhoDFQAtM12Wjo2xxA5sduzB/3HdzZmiSKCBgzCBgKR+MHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6MewAjAiGA8y
# MDIzMTAwNDA5MzM1NFoYDzIwMjMxMDA1MDkzMzU0WjB0MDoGCisGAQQBhFkKBAEx
# LDAqMAoCBQDox7ACAgEAMAcCAQACAhqzMAcCAQACAhRRMAoCBQDoyQGCAgEAMDYG
# CisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEA
# AgMBhqAwDQYJKoZIhvcNAQELBQADggEBADO2WIrt3gvA6IILC4HRc7AzzsnT6k+m
# OaXwlaR1TpgybrNAqUVusdO15muhtK/+lmC9CiJDl2jkhIs6RaP9BWbKGiMFSU3n
# 1r0q0V7iIKaT87tstqHlEqUSU4ONS8mUmxGxyO0laRx2IE1+xUqZSM6dYizFm473
# eXjqXECE8Atq39UikguAj9piVmTNcaB57D03PoY4Ts0S1UdEiLMGlICfx4cyyfEN
# vwl3EeRsa1i4h6lmjKF66uExzBro7GjV9tvLsUkR2YryvTVX0gaqnd3To/FF8Nvu
# GiRwLjmWojS0aeeCW2lnSQkjs2lc3ob3HUUgSBJ8B8m74lRQgsjkHF0xggQNMIIE
# CQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYw
# JAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdTk6QMv
# wKxprAABAAAB1DANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqG
# SIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCBES41dOJhXAOKEiC4Kq8DQS22wy5k5
# d7NnqVUR1t78RDCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIMzqh/rYFKXO
# lzvWS5xCtPi9aU+fBUkxIriXp2WTPWI3MIGYMIGApH4wfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAHU5OkDL8CsaawAAQAAAdQwIgQgiiCJkJbHgNzy
# Hmi/ZzMf46uIrdLGR6BXdpHe96WZVtEwDQYJKoZIhvcNAQELBQAEggIAhELG1LIX
# zHl4yj6YBZNvNX9YuI9XrSUVuyrB3mcuxH1BO8ylMXYOp/KJwuvlsyP1SaUrppz1
# N+uom5w/qe/EorVKzZAAaMFqmAW+SLYxXB/Fw9/IyNh6C3jhCcUmtX+fbOJrnwyy
# fSO94bbc/IJCpogENBXbqa+pGSbFosbcJZX2LlOVAvGTgL7CVk43SI5X2nWHFvrX
# TrSJ5ePzr4ykRjF8K9XL0GvOXRMieybOfXhlRrzrlUjYvL7TNgVPUUPnhx50E1dl
# coNVXIrq34pv5dKNucb9z3jiJvU4GsYDYOx9dx8ojlqwLzPE+8fVyy12wU3yIxKY
# uNazKMewdF1tBUro0NWYo5ebW0TX/4F71etZeXpkYKW87IWxUCtD6mjRMqwlL/GV
# t0AZ+cligALt4Lb89o1jiKHnFVSNq30PbZZz2bo5rKFx/TqJ8oN0Y4JG9YnV7Dyy
# WgwgPa1e8i6e5dAES3BEKvpAOvS5UjYCPpwzZVBzjCWEsUFx3m2rNwNPK0WlUpjJ
# W5UVuoHIkHQ3kbX7ca5MoLNIyfkizn+9V34tMkD9OICCUQiy7RVR6w1BvbAy8doQ
# 7647/FkFBxTRhDSVu+Sn1q2Yg6IsyxjzuilmwjOM+AusFbkrQBJo4gHdIGAAIp04
# k2k0wDttCzG5zzVtdxDfo0FCdrbQw/lldtY=
# SIG # End signature block
