#-------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation.  All rights reserved.
#-------------------------------------------------------------------------

Set-Variable ObjectDiagnosticsReportOutputDirectory          "$env:ProgramData\AADConnect\ADSyncObjectDiagnostics"                                                     -Option Constant -Scope global

Set-Variable SynchronizationIssueHtmlItems                   (New-Object System.Collections.Generic.List[string])                                                      -Option Constant -Scope script

Set-Variable HtmlGroupList                                   (New-Object System.Collections.Generic.List[string])                                                      -Option Constant -Scope script

Set-Variable SynchronizationIssueList                        (New-Object System.Collections.Generic.List[string])                                                      -Option Constant -Scope script

#
# Event IDs
#
Set-Variable EventIdSingleObjectTroubleshootingRun           2501                                                                                                      -Option Constant -Scope script
Set-Variable EventIdMVObjectAttributeNotFound                2502                                                                                                      -Option Constant -Scope script
Set-Variable EventIdFailedToImportMSOnlineModule             2503                                                                                                      -Option Constant -Scope script
Set-Variable EventIdDiagnoseSingleObject                     2504                                                                                                      -Option Constant -Scope script
Set-Variable EventIdUPNMismatch                              2505                                                                                                      -Option Constant -Scope script
Set-Variable EventIdUPNUpdateBlocked                         2506                                                                                                      -Option Constant -Scope script
Set-Variable EventIdUPNSuffixNotVerified                     2507                                                                                                      -Option Constant -Scope script
Set-Variable EventIdUPNSuffixVerified                        2508                                                                                                      -Option Constant -Scope script
Set-Variable EventIdUPNFederatedDomainChange                 2509                                                                                                      -Option Constant -Scope script
Set-Variable EventIdDomainFiltered                           2510                                                                                                      -Option Constant -Scope script
Set-Variable EventIdOUFiltered                               2511                                                                                                      -Option Constant -Scope script
Set-Variable EventIdCsObjectAttributeNotFound                2512                                                                                                      -Option Constant -Scope script
Set-Variable EventIdDynamicDistributionGroup                 2513                                                                                                      -Option Constant -Scope script
Set-Variable EventIdLinkedMailboxIssue                       2514                                                                                                      -Option Constant -Scope script
Set-Variable EventIdGroupMembershipTroubleshootingRun        2515                                                                                                      -Option Constant -Scope script
Set-Variable EventIdIsToolHelpful                            2516                                                                                                      -Option Constant -Scope script
Set-Variable EventIdConnectorAccountReadPermissions          2517                                                                                                      -Option Constant -Scope script
Set-Variable EventIdCloudOwnedAttributes                     2518                                                                                                      -Option Constant -Scope script
Set-Variable EventIdObjectTypeInclusion                      2519                                                                                                      -Option Constant -Scope script
Set-Variable EventIdGroupFiltered                            2520                                                                                                      -Option Constant -Scope script

#
# Event Messages
#
Set-Variable EventMsgSingleObjectTroubleshootingRun          "Single object troubleshooting workflow has been run."                                                    -Option Constant -Scope script
Set-Variable EventMsgGroupMembershipTroubleshootingRun       "Group membership troubleshooting workflow has been run."                                                    -Option Constant -Scope script
Set-Variable EventMsgMVObjectAttributeNotFound               "Metaverse object attribute is not found. Attribute Name: {0}."                                           -Option Constant -Scope script
Set-Variable EventMsgFailedToImportMSOnlineModule            "Failed to import MSOnline Module."                                                                       -Option Constant -Scope script
Set-Variable EventMsgDiagnoseSingleObject                    "Single object diagnostics has been run."                                                                   -Option Constant -Scope script
Set-Variable EventMsgUPNMismatch                             "UPN Mismatch. AADConnect Object UPN: {0}, AAD Tenant Object UPN: {1}."                                   -Option Constant -Scope script
Set-Variable EventMsgUPNUpdateBlocked                        "UPN update is blocked for the AAD Tenant Object: {0}."                                                   -Option Constant -Scope script
Set-Variable EventMsgUPNSuffixNotVerified                    "AADConnect object UPN suffix {0} is NOT verified with AAD Tenant {1}."                                   -Option Constant -Scope script
Set-Variable EventMsgUPNSuffixVerified                       "AADConnect object UPN suffix {0} is verified with AAD Tenant {1}."                                       -Option Constant -Scope script
Set-Variable EventMsgUPNFederatedDomainChange                "Federated Domain Change. AADConnect Object UPN: {0}, AAD Tenant Object UPN: {1}."                        -Option Constant -Scope script
Set-Variable EventMsgDomainFiltered                          "Object {0} filtered due to domain filtering. Domain: {1}"                                                -Option Constant -Scope script
Set-Variable EventMsgOUFiltered                              "Object {0} filtered due to OU filtering. OU: {1}"                                                        -Option Constant -Scope script
Set-Variable EventMsgCsObjectAttributeNotFound               "Connector space object attribute is not found. Attribute Name: {0}."                                     -Option Constant -Scope script
Set-Variable EventMsgDynamicDistributionGroup                "Object is not synchronized since it is a dynamic distribution group."                                    -Option Constant -Scope script
Set-Variable EventMsgLinkedMailboxIssue                      "Object is not synchronized since it has on-premises linked mailbox."                                     -Option Constant -Scope script
Set-Variable EventMsgConnectorAccountReadPermissions         "Comparing read permissions on object. AD Connector Name: {0}, Object DN: {1}, Connector account: {2}, Provided account: {3}"     -Option Constant -Scope script
Set-Variable EventMsgCloudOwnedAttributes                    "Cloud owned attributes found. AAD Object DN: {0}, Attributes: {1}"                                       -Option Constant -Scope script
Set-Variable EventMsgObjectTypeInclusion                     "Object {0} is of type '{1}' which is not part of the object type inclusion list for connector {2}"       -Option Constant -Scope script
Set-Variable EventMsgGroupFiltered                           "Object {0} filtered due to group filtering. Connector: {1}, Group filtering group: {2}"                  -Option Constant -Scope script

Function Debug-ADSyncObjectSynchronizationIssuesNonInteractiveMode
{
    param
    (
        [string]
        [parameter(mandatory=$false)]
        $ADConnectorName,

        [string]
        [parameter(mandatory=$false)]
        $ObjectDN,

        [string]
        [parameter(mandatory=$true)]
        $DiagnosticOption
    )

    $timezone = [TimeZoneInfo]::Local
    ReportOutput -PropertyName 'Sync Server TimeZone' -PropertyValue $timezone.DisplayName

    Debug-ADSyncObjectSynchronizationIssues -ADConnectorName $ADConnectorName -ObjectDN $ObjectDN -DiagnosticOption $DiagnosticOption
}

Function Debug-ADSyncGroupMembershipSynchronizationIssuesNonInteractiveMode
{
    param
    (
        [string]
        [parameter(mandatory=$false)]
        $GroupADConnectorName,

        [string]
        [parameter(mandatory=$false)]
        $GroupDN,

        [string]
        [parameter(mandatory=$false)]
        $MemberADConnectorName,

        [string]
        [parameter(mandatory=$false)]
        $MemberDN
    )

    $timezone = [TimeZoneInfo]::Local
    ReportOutput -PropertyName 'Sync Server TimeZone' -PropertyValue $timezone.DisplayName

    Debug-ADSyncGroupMembershipSynchronizationIssues -GroupADConnectorName $GroupADConnectorName -GroupDN $GroupDN -MemberADConnectorName $MemberADConnectorName -MemberDN $MemberDN
}

Function Debug-ADSyncAttributeSynchronizationIssues
{
    param
    (
        [string]
        [parameter(mandatory=$false)]
        $ADConnectorName,

        [string]
        [parameter(mandatory=$false)]
        $ObjectDN,

        [string]
        [parameter(mandatory=$true)]
        $DiagnosticOption
    )

    WriteEventLog($EventIdSingleObjectTroubleshootingRun)($EventMsgSingleObjectTroubleshootingRun)

    Write-Host "`r`n"

    $title = $DiagnosticOption

    #
    # Write Title to the PowerShell Console
    #
    WriteTitle($title)

    try
    {
        $aadConnectRegKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Azure AD Connect'
        $aadConnectWizardPath = $aadConnectRegKey.WizardPath
        $aadConnectWizardFileName = "AzureADConnect.exe"

        $aadConnectPathLength = $aadConnectWizardPath.IndexOf($aadConnectWizardFileName)
        $aadConnectPath = $aadConnectWizardPath.Substring(0, $aadConnectPathLength)

        $modulePath    = [System.IO.Path]::Combine($aadConnectPath, "AADPowerShell\MSOnline.psd1")
        Import-Module $modulePath -ErrorAction Stop

        $adSyncToolsModulePath = [System.IO.Path]::Combine($aadConnectPath, "Tools\AdSyncTools.psm1")
        Import-Module $adSyncToolsModulePath -ErrorAction Stop
    }
    catch
    {
        WriteEventLog($EventIdFailedToImportMSOnlineModule)($EventMsgFailedToImportMSOnlineModule)

        "MSOnline module needs to be imported" | ReportError
        return
    }

    Write-Host "`r`n"

    $adConnector = GetADConnectorByName($ADConnectorName)("Please enter AD Connector Name")
    if ($adConnector -eq $null)
    {
        "There is no AD Connector with name `"$ADConnectorName`"." | ReportError
        Write-Host "`r`n"
        return
    }

    $ADConnectorName = $adConnector.Name
    $ObjectDN = GetObjectDN($ObjectDN)("Please enter AD object Distinguished Name")

    #
    # Validate ObjectDN parameter
    #
    $adObject = Search-ADSyncDirectoryObjects -AdConnectorId $adConnector.Identifier -LdapFilter "(distinguishedName=$ObjectDN)" -SearchScope Subtree -SizeLimit 1
    if ($adObject -eq $null -or $adObject.Count -lt 1)
    {
        "Could not find an object in on-premises AD with distinguishedName=`"$ObjectDN`"." | ReportError
        
        Write-Host "`r`n"
        return
    }

    #
    # Check if object is a dynamic distribution group
    #
    $isDynamicDistributionGroup = IsObjectTypeMatch($adObject[0])("msExchDynamicDistributionList")
    if ($isDynamicDistributionGroup -eq $true)
    {
        WriteEventLog($EventIdDynamicDistributionGroup)($EventMsgDynamicDistributionGroup)

        $SynchronizationIssueList.Add($global:DynamicDistributionGroupIssue)

        "The given object is a dynamic distribution group. Azure AD Connect does not synchronize on-premises dynamic distribution groups to Azure AD." | ReportError -PropertyName "Dynamic Distribution Group" -PropertyValue "True"
        Write-Host "`r`n"

        AskIfToolHelpful($ObjectDN)

        $SynchronizationIssueList.Clear()
        return
    }
    else
    {
        ReportOutput -PropertyName "Dynamic Distribution Group" -PropertyValue "False"
    }

    #
    # Get object graph
    #
    $adCsObject = $null
    $mvObject = $null
    $aadCsObject = $null
    $objectGraph = CheckObjectGraph($ADConnectorName)($ObjectDN)([ref]$adCsObject)([ref]$mvObject)([ref]$aadCsObject)	

    $cloudownedAttributes = $null

    foreach ($attribute in $aadCsObject.Attributes)
    {
        if ($attribute.Name -ieq 'cloudMasteredProperties')
        {
            foreach ($value in $attribute.Values)
            {
                if ($cloudownedAttributes -eq $null)
                {
                    $cloudownedAttributes = "[" + $value
                }
                else
                {
                    $cloudownedAttributes += " ," + $value
                }
            }
            if ($cloudownedAttributes -ne $null)
            {
                $cloudownedAttributes += "]"
            }
        }
    }

    if ($cloudownedAttributes -ne $null)
    {
        $SynchronizationIssueList.Add($global:CloudOwnedAttributeIssue)
        WriteEventLog($EventIdCloudOwnedAttributes)($EventMsgCloudOwnedAttributes -f ($aadCsObject.DistinguishedName, $cloudownedAttributes))
        "The given AD CS Object:`r`n" + $adCsObject.DistinguishedName + "`r`nAAD CS object:`r`n" + $aadCsObject.DistinguishedName + "`r`nhas cloud owned attributes:`r`n" + $cloudownedAttributes + "`r`nAzure AD Connect does not synchronize cloud owned attributes to Azure AD." | ReportError -PropertyName "Cloud owned attributes found" -PropertyValue "True"
        Write-Host "`r`n"

        AskIfToolHelpful($ObjectDN)

        $SynchronizationIssueList.Clear()
        return
    }
    else
    {
        ReportOutput -PropertyName "Cloud owned attributes found" -PropertyValue "False"
    }


    #
    # Ask if the tool is helpful for the synchronization issues
    #
    # The customer is going to answer for each synchronization issue separately
    #
    AskIfToolHelpful($ObjectDN)

    $SynchronizationIssueList.Clear()
}
Function Debug-ADSyncObjectSynchronizationIssues
{
    param
    (
        [string]
        [parameter(mandatory=$false)]
        $ADConnectorName,

        [string]
        [parameter(mandatory=$false)]
        $ObjectDN,

        [string]
        [parameter(mandatory=$true)]
        $DiagnosticOption
    )

    WriteEventLog($EventIdSingleObjectTroubleshootingRun)($EventMsgSingleObjectTroubleshootingRun)

    Write-Host "`r`n"

    $title = $DiagnosticOption

    #
    # Write Title to the PowerShell Console
    #
    WriteTitle($title)

    try
    {
        $aadConnectRegKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Azure AD Connect'
        $aadConnectWizardPath = $aadConnectRegKey.WizardPath
        $aadConnectWizardFileName = "AzureADConnect.exe"

        $aadConnectPathLength = $aadConnectWizardPath.IndexOf($aadConnectWizardFileName)
        $aadConnectPath = $aadConnectWizardPath.Substring(0, $aadConnectPathLength)

        $modulePath    = [System.IO.Path]::Combine($aadConnectPath, "AADPowerShell\MSOnline.psd1")
        Import-Module $modulePath -ErrorAction Stop

        $adSyncToolsModulePath = [System.IO.Path]::Combine($aadConnectPath, "Tools\AdSyncTools.psm1")
        Import-Module $adSyncToolsModulePath -ErrorAction Stop
    }
    catch
    {
        WriteEventLog($EventIdFailedToImportMSOnlineModule)($EventMsgFailedToImportMSOnlineModule)

        "MSOnline module needs to be imported" | ReportError
        return
    }

    Write-Host "`r`n"

    $adConnector = GetADConnectorByName($ADConnectorName)("Please enter AD Connector Name")
    if ($adConnector -eq $null)
    {
        "There is no AD Connector with name `"$ADConnectorName`"." | ReportError
        Write-Host "`r`n"
        return
    }

    $ADConnectorName = $adConnector.Name
    $ObjectDN = GetObjectDN($ObjectDN)("Please enter AD object Distinguished Name")

    Write-Host "`r`n"
    Write-Host "Searching for object `"$ObjectDN`" using `"$($adConnector.Name)`" Connector credentials `"$($adConnector.ConnectivityParameters["forest-login-domain"].Value)\$($adConnector.ConnectivityParameters["forest-login-user"].Value)`"..."
    Write-Host "`r`n"

    #
    # Validate ObjectDN parameter
    #
    $adObject = Search-ADSyncDirectoryObjects -AdConnectorId $adConnector.Identifier -LdapFilter "(distinguishedName=$ObjectDN)" -SearchScope Subtree -SizeLimit 1
    if ($adObject -eq $null -or $adObject.Count -lt 1)
    {
        "Could not find an object in on-premises AD with distinguishedName=`"$ObjectDN`" using Connector credentials." | ReportError

        if (!$isNonInteractiveMode)
        {
            $searchWithAdminOptions = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes", "&No")
            $searchWithAdmin = !($host.UI.PromptForChoice("Confirm object existence", "Would you like to search for the object using admin credentials?", $searchWithAdminOptions, 0))

            if ($searchWithAdmin)
            {
                $ADForest = $adConnector.ConnectivityParameters["forest-name"].Value
                $ADForestCredential = GetValidADForestCredentials($ADForest)
                Write-Host "`r`n"

                if ($ADForestCredential -eq $null)
                {
                    "No valid credentials were provided for forest with name=`"$ADForest`"." | ReportError
                }
                else
                {
                    $adObjectFromProvidedAccount = Search-ADSyncDirectoryObjects -AdConnectorId $adConnector.Identifier -LdapFilter "(distinguishedName=$ObjectDN)" -SearchScope Subtree -SizeLimit 1 -AdConnectorCredential $ADForestCredential -ForestFqdn $ADForest

                    if ($adObjectFromProvidedAccount -eq $null -or $adObjectFromProvidedAccount.Count -lt 1)
                    {
                        "Could not find an object in on-premises AD with distinguishedName=`"$ObjectDN`" using provided credentials. Please ensure the object exits in AD and the Distinguished Name is correct." | ReportError
                    }
                    else
                    {
                        "Object found in forest using admin credentials. Connector credentials to not have sufficient permissions to read the object. More information on configuring permissions can be found here: `"$global:ConfigureAccountPermissionsUrl`"" | ReportError
                    }
                }
            }
        }

        Write-Host "`r`n"
        return
    }

    #
    # Check if object is a dynamic distribution group
    #
    $isDynamicDistributionGroup = IsObjectTypeMatch($adObject[0])("msExchDynamicDistributionList")
    if ($isDynamicDistributionGroup -eq $true)
    {
        WriteEventLog($EventIdDynamicDistributionGroup)($EventMsgDynamicDistributionGroup)

        $SynchronizationIssueList.Add($global:DynamicDistributionGroupIssue)

        "The given object is a dynamic distribution group. Azure AD Connect does not synchronize on-premises dynamic distribution groups to Azure AD." | ReportError -PropertyName "Dynamic Distribution Group" -PropertyValue "True"
        Write-Host "`r`n"

        AskIfToolHelpful($ObjectDN)

        $SynchronizationIssueList.Clear()
        return
    }
    else
    {
        ReportOutput -PropertyName "Dynamic Distribution Group" -PropertyValue "False"
    }

    #
    # Active Directory Connector Partition Details
    #
    $partitionDN = Get-DirectoryPartitionDN($ObjectDN)
    $partition = Get-ADSyncConnectorPartition -Connector $adConnector -Name $partitionDN    

    #
    # Check Domain/OU Filtering, object type inclusion, and group filtering
    #
    CheckDomainBasedFiltering($partition)($adConnector)($ObjectDN)
    CheckOUBasedFiltering($partition)($ObjectDN)
    CheckObjectTypeInclusion($adConnector)($adObject[0])
    CheckGroupFiltering($adConnector)($adObject[0])

    #
    # Get object graph
    #
    $adCsObject = $null
    $mvObject = $null
    $aadCsObject = $null
    $objectGraph = CheckObjectGraph($ADConnectorName)($ObjectDN)([ref]$adCsObject)([ref]$mvObject)([ref]$aadCsObject)	

    if ($adCsObject -ne $null)
    {
        #
        # Check attribute based filtering
        #
        CheckAttributeBasedFiltering($ADConnectorName)($ObjectDN)($adCsObject)($mvObject)($aadCsObject)

        # Check for Sync Errors and Export Errors
        if ((ReportSyncError($adCsObject)) -or (ReportSyncError($aadCsObject)))
        {
            ReportOutput -PropertyName "Sync Error(s)" -PropertyValue "True"
        }

        if ((ReportExportError($adCsObject)) -or (ReportExportError($aadCsObject)))
        {
            ReportOutput -PropertyName "Export Error(s)" -PropertyValue "True"
        }	

        # Check for transient
        if ((ReportTransient($adCsObject)) -or (ReportTransient($aadCsObject)))
        {
            ReportOutput -PropertyName "Transient Object" -PropertyValue "True"
        }

        #
        # Connect to Azure Active Directory
        #
        ConnectToAAD

        #
        # Get all domains registered to AAD Tenant.
        #
        # Each AAD Tenant includes 2 default verified domains.
        #
        #    1. <Initial-default-domain-name>.onmicrosoft.com
        #    2. <Initial-default-domain-name>.mail.onmicrosoft.com
        #
        # Each AAD Tenant is associated with an initial default domain name.
        #
        $aadTenantDomains = GetAADDomains
    
        #
        # Get default domain name "<Initial-default-domain-name>.onmicrosoft.com" which replaces
        # non-routable OR unverified OR notAdded on-premises upn sufixes.
        #
        $aadTenantDefaultDomainName = GetAADTenantDefaultDomainName($aadTenantDomains)
    }

    #
    # Get AAD Tenant object and create html group
    #
    if (IsObjectTypeMatch($adObject[0])("user"))
    {
        $aadTenantUser = $null
        if ($mvObject -and $aadCsObject)
        {
             $aadTenantUser = GetAADTenantUser($mvObject)
        }

        $userObjectDetailsHtmlGroup = GetUserObjectHtmlGroup($adObject[0])($adCsObject)($mvObject)($aadCsObject)($aadTenantUser)
        $HtmlGroupList.Insert(0, $userObjectDetailsHtmlGroup)
    }
    elseif (IsObjectTypeMatch($adObject[0])("group"))
    {
        $aadTenantGroup = $null
        if ($mvObject -and $aadCsObject)
        {
             $aadTenantGroup = GetAADTenantGroup($mvObject)
        }

        $groupObjectDetailsHtmlGroup = GetGroupObjectHtmlGroup($adObject[0])($adCsObject)($mvObject)($aadCsObject)($aadTenantGroup)
        $HtmlGroupList.Insert(0, $groupObjectDetailsHtmlGroup)
    }
    elseif (IsObjectTypeMatch($adObject[0])("contact"))
    {
        $aadTenantContact = $null
        if ($mvObject -and $aadCsObject)
        {
             $aadTenantContact = GetAADTenantContact($mvObject)
        }

        $contactObjectDetailsHtmlGroup = GetContactObjectHtmlGroup($adObject[0])($adCsObject)($mvObject)($aadCsObject)($aadTenantContact)
        $HtmlGroupList.Insert(0, $contactObjectDetailsHtmlGroup)
    }

    #
    # Diagnose Issues
    #
    if ($title -eq $global:DiagnoseObjectSyncIssues)
    {
        if ($adCsObject -ne $null)
        {
            if (IsObjectTypeMatch($adObject[0])("user"))
            {
                DiagnoseUserObject($adConnector)($adObject[0])($adCsObject)($mvObject)($aadCsObject)($aadTenantUser)($aadTenantDefaultDomainName)($aadTenantDomains)
            }
            elseif(IsObjectTypeMatch($adObject[0])("group"))
            {
                DiagnoseGroupObject($adConnector)($adObject[0])($adCsObject)($mvObject)($aadCsObject)($aadTenantGroup)
            }
            elseif(IsObjectTypeMatch($adObject[0])("foreignSecurityPrincipal"))
            {
                DiagnoseFSP($adConnector)($adObject[0])($adCsObject)($mvObject)($aadCsObject)
            }
            elseif (IsObjectTypeMatch($adObject[0])("contact"))
            {
            }
            else
            {
                "This diagnostic option is currently only supported on user, contact, group and foreignSecurityPrincipal objects." | ReportError
                Write-Host "`r`n"
            }
        }

        if ($SynchronizationIssueHtmlItems.Count -eq 0)
        {
            Write-Host "`r`n"
            "We couldn't identify the issue with this object, please contact customer support if you need assistance, or read the AADConnect documentation at `"$global:TroubleshootingTaskUrl`"" | ReportWarning
        }
    }
    elseif ($title -eq $global:ChangePrimaryEmailAddress)
    {
        if (IsObjectTypeMatch($adObject[0])("user"))
        {
            DiagnoseUserObjectProxyAddresses
        }
        else
        {
            "This diagnostic option is currently only supported on user objects." | ReportError
            Write-Host "`r`n"
        }
    }
    elseif ($title -eq $global:HideFromGlobalAddressList)
    {
        if ($adCsObject -ne $null)
        {
            if (IsObjectTypeMatch($adObject[0])("user"))
            {
                DiagnoseUserObjectHideFromAddressLists($AdConnector)($AdObject[0])($AdCsObject)
            }
            else
            {
                "This diagnostic option is currently only supported on user objects." | ReportError
                Write-Host "`r`n"
            }
        }
    }

    #
    # Set output directory for per-object html report
    #
    Set-OutputDirectory

    #
    # Per-object html report date-time
    #
    $reportDate = [string] $(Get-Date -Format yyyyMMddHHmmss)

    
    if ($HtmlGroupList.Count -gt 0)
    {
        if ($SynchronizationIssueHtmlItems.Count -gt 0)
        {
            $synchronizationIssuesHtmlGroup = WriteHtmlAccordionGroup($SynchronizationIssueHtmlItems)($global:HtmlSynchronizationIssuesSectionTitle)

            $HtmlGroupList.Insert(0, $synchronizationIssuesHtmlGroup)
        }

        $objectSyncDiagnosticsHtmlContent = WriteHtmlAccordion($HtmlGroupList)($ObjectDN)

        $objectSyncDiagnosticsHtmlBody = WriteHtmlBody($objectSyncDiagnosticsHtmlContent)

        $objectSyncDiagnosticsHtml = WriteHtml($objectSyncDiagnosticsHtmlBody)

        Write-Host "`r`n"

        Export-ObjectDiagnosticsHtmlReport -Title $ObjectDN -ReportDate $reportDate -HtmlDoc $objectSyncDiagnosticsHtml

        $SynchronizationIssueHtmlItems.Clear()
        $HtmlGroupList.Clear()
    }

    Write-Host "`r`n"
    Write-Host "`r`n"

    #
    # Ask if the tool is helpful for the synchronization issues
    #
    # The customer is going to answer for each synchronization issue separately
    #
    AskIfToolHelpful($ObjectDN)

    $SynchronizationIssueList.Clear()
}

Function Debug-ADSyncObjectAttributeRetrievalIssues
{
    param
    (
        [string]
        [parameter(mandatory=$false)]
        $ADConnectorName,

        [string]
        [parameter(mandatory=$false)]
        $ObjectDN
    )

    if ($isNonInteractiveMode)
    {
        return
    }

    Write-Host "`r`n"

    #
    # Write Title to the PowerShell Console
    #
    WriteTitle($global:ADConnectorAccountReadPermissions)

    $adConnector = GetADConnectorByName($ADConnectorName)("Please enter AD Connector Name")
    if ($adConnector -eq $null)
    {
        "There is no AD Connector with name `"$ADConnectorName`"." | ReportError
        Write-Host "`r`n"
        return
    }

    $ObjectDN = GetObjectDN($ObjectDN)("Please enter AD object Distinguished Name")
    $adObject = Search-ADSyncDirectoryObjects -AdConnectorId $adConnector.Identifier -LdapFilter "(distinguishedName=$ObjectDN)" -SearchScope Subtree -SizeLimit 1
    if ($adObject -eq $null -or $adObject.Count -lt 1)
    {
        "Could not find an object in on-premises AD with distinguishedName=`"$ObjectDN`"." | ReportError
        Write-Host "`r`n"
        return
    }

    $ADForest = $adConnector.ConnectivityParameters["forest-name"].Value
    $ADConnectorAccountName = "$($adConnector.ConnectivityParameters['forest-login-domain'].Value)\$($adConnector.ConnectivityParameters['forest-login-user'].Value)"

    $ADForestCredential = GetValidADForestCredentials($ADForest)
    if ($ADForestCredential -eq $null)
    {
        "No valid credentials were provided for forest with name=`"$ADForest`"." | ReportError
        Write-Host "`r`n"
        return
    }

    WriteEventLog($EventIdConnectorAccountReadPermissions)($EventMsgConnectorAccountReadPermissions -f ($ADConnectorName, $ObjectDN, $ADConnectorAccountName, $ADForestCredential.UserName))

    $adObjectFromProvidedAccount = Search-ADSyncDirectoryObjects -AdConnectorId $adConnector.Identifier -LdapFilter "(distinguishedName=$ObjectDN)" -SearchScope Subtree -SizeLimit 1 -AdConnectorCredential $ADForestCredential -ForestFqdn $ADForest

    $objectAttributesFromConnectorAccount = [System.Collections.Generic.Dictionary[[String], [Object]]] $adObject[0]
    $objectAttributesFromProvidedAccount = [System.Collections.Generic.Dictionary[[String], [Object]]] $adObjectFromProvidedAccount[0]

    $attributeDetailsHtmlGroup = GetObjectAllADAttributeHtmlGroup($objectAttributesFromConnectorAccount)($objectAttributesFromProvidedAccount)($ADConnectorAccountName)($ADForestCredential.UserName)

    $HtmlGroupList.Insert(0, $attributeDetailsHtmlGroup)

    #
    # Set output directory for per-object html report
    #
    Set-OutputDirectory

    #
    # Per-object html report date-time
    #
    $reportDate = [string] $(Get-Date -Format yyyyMMddHHmmss)
    

    $objectSyncDiagnosticsHtmlContent = WriteHtmlAccordion($HtmlGroupList)($ObjectDN)

    $objectSyncDiagnosticsHtmlBody = WriteHtmlBody($objectSyncDiagnosticsHtmlContent)

    $objectSyncDiagnosticsHtml = WriteHtml($objectSyncDiagnosticsHtmlBody)

    Write-Host "`r`n"

    Export-ObjectDiagnosticsHtmlReport -Title $ObjectDN -ReportDate $reportDate -HtmlDoc $objectSyncDiagnosticsHtml

    $HtmlGroupList.Clear()
    

    Write-Host "`r`n"
    Write-Host "`r`n"
}

Function Debug-ADSyncGroupMembershipSynchronizationIssues
{
    param
    (
        [string]
        [parameter(mandatory=$false)]
        $GroupADConnectorName,

        [string]
        [parameter(mandatory=$false)]
        $GroupDN,

        [string]
        [parameter(mandatory=$false)]
        $MemberADConnectorName,

        [string]
        [parameter(mandatory=$false)]
        $MemberDN
    )

    WriteEventLog($EventIdGroupMembershipTroubleshootingRun)($EventMsgGroupMembershipTroubleshootingRun)

    Write-Host "`r`n"

    #
    # Write Title to the PowerShell Console
    #
    WriteTitle($global:DiagnoseGroupMembershipSyncIssues)    

    #
    # Get AD Connector Name for the group
    #
    $adConnectorGroup = GetADConnectorByName($GroupADConnectorName)("Please enter AD Connector Name for the group")
    if ($adConnectorGroup -eq $null)
    {
        "There is no AD Connector with name `"$GroupADConnectorName`"." | ReportError
        Write-Host "`r`n"
        return
    }

    $GroupADConnectorName = $adConnectorGroup.Name
    $GroupDN = GetObjectDN($GroupDN)("Please enter group Distinguished Name")
    
    #
    # Validate GroupDN parameter
    #
    $adGroupObject = Search-ADSyncDirectoryObjects -AdConnectorId $adConnectorGroup.Identifier -LdapFilter "(distinguishedName=$GroupDN)" -SearchScope Subtree -SizeLimit 1
    if ($adGroupObject -eq $null -or $adGroupObject.Count -lt 1)
    {
        "Could not find an object in on-premises AD with distinguishedName=`"$GroupDN`"." | ReportError        
        Write-Host "`r`n"
        return
    }

    #
    # Validating that the object is a group
    #
    $isGroup = IsObjectTypeMatch($adGroupObject[0])("group")
    if ($isGroup -eq $false)
    {
        "Given object with distinguishedName=`"$GroupDN`" is not a group." | ReportError
        Write-Host "`r`n"
        return
    }	

    #
    # Active Directory Connector Partition Details for group
    #
    $partitionDNGroup = Get-DirectoryPartitionDN($GroupDN)
    $partitionGroup = Get-ADSyncConnectorPartition -Connector $adConnectorGroup -Name $partitionDNGroup

    #
    # Check Domain/OU Filtering for group
    #
    CheckDomainBasedFiltering($partitionGroup)($adConnectorGroup)($GroupDN)
    CheckOUBasedFiltering($partitionGroup)($GroupDN)

    #
    # Get AD Connector Name for the group member
    #
    $adConnectorMember = GetADConnectorByName($MemberADConnectorName)("Please enter AD Connector Name for the group member")
    if ($adConnectorMember -eq $null)
    {
        "There is no AD Connector with name `"$MemberADConnectorName`"." | ReportError
        Write-Host "`r`n"
        return
    }

    $MemberADConnectorName = $adConnectorMember.Name
    $MemberDN = GetObjectDN($MemberDN)("Please enter Distinguished Name for the group member")
    
    #
    # Validate MemberDN parameter
    #
    $adMemberObject = Search-ADSyncDirectoryObjects -AdConnectorId $adConnectorMember.Identifier -LdapFilter "(distinguishedName=$MemberDN)" -SearchScope Subtree -SizeLimit 1
    if ($adMemberObject -eq $null -or $adMemberObject.Count -lt 1)
    {
        "Could not find an object in on-premises AD with distinguishedName=`"$MemberDN`"." | ReportError        
        Write-Host "`r`n"
        return
    }	 		

    #
    # Active Directory Connector Partition Details for group member
    #
    $partitionDNMember = Get-DirectoryPartitionDN($MemberDN)
    $partitionMember = Get-ADSyncConnectorPartition -Connector $adConnectorMember -Name $partitionDNMember

    #
    # Check Domain/OU Filtering for group member
    #
    CheckDomainBasedFiltering($partitionMember)($adConnectorMember)($MemberDN)    
    CheckOUBasedFiltering($partitionMember)($MemberDN)

    #
    # Foreign Security Principal
    #
    $isForeignSecurityPrincipal = $false
    if ($GroupADConnectorName -ne $MemberADConnectorName)
    {
        $isForeignSecurityPrincipal = $true
        $memberSidBinary = $adMemberObject[0]["objectsid"]
        $memberSidString = (New-Object System.Security.Principal.SecurityIdentifier($memberSidBinary[0], 0)).Value

        $MemberDN = "CN=$memberSidString,CN=ForeignSecurityPrincipals,$partitionDNGroup"

        $adMemberObject = Search-ADSyncDirectoryObjects -AdConnectorId $adConnectorGroup.Identifier -LdapFilter "(distinguishedName=$MemberDN)" -SearchScope Subtree -SizeLimit 1
        if ($adMemberObject -eq $null -or $adMemberObject.Count -lt 1)
        {
            "Could not find an object in on-premises AD with distinguishedName=`"$MemberDN`"." | ReportError
            Write-Host "`r`n"
            return
        }
        else
        {
            "Foreign Security Principal `"$memberSidString`" exists in Active Directory" | ReportOutput
            Write-Host "`r`n"
        }

        #
        # Check OU Filtering for group member which is a foreign security principal
        #		
        CheckOUBasedFiltering($partitionGroup)($MemberDN)
    }

    # Validate if member reference is present in AD
    $isMemberInAD = IsMemberOfGroupInAD($adGroupObject[0])($MemberDN)
    if (-not $isMemberInAD)
    {
        "Object `"$MemberDN`" is not a member of group `"$GroupDN`" in Active Directory" | ReportError
        Write-Host "`r`n"
        return
    }

    #
    # Get object graph for group
    #
    $adCsObjectGroup = $null
    $mvObjectGroup = $null
    $aadCsObjectGroup = $null
    $objectGraphGroup = CheckObjectGraph($GroupADConnectorName)($GroupDN)([ref]$adCsObjectGroup)([ref]$mvObjectGroup)([ref]$aadCsObjectGroup)	

    #
    # Check attribute based filtering for group
    #
    CheckAttributeBasedFiltering($GroupADConnectorName)($GroupDN)($adCsObjectGroup)($mvObjectGroup)($aadCsObjectGroup)
    
    # Check for Sync Errors and Export Errors on group object
    if ((ReportSyncError($adCsObjectGroup)) -or (ReportSyncError($aadCsObjectGroup)))
    {
        ReportOutput -PropertyName "Sync Error(s) on group" -PropertyValue "True"
    }

    if ((ReportExportError($adCsObjectGroup)) -or (ReportExportError($aadCsObjectGroup)))
    {
        ReportOutput -PropertyName "Export Error(s) on group" -PropertyValue "True"
    }	

    # Check for transient on group object
    if ((ReportTransient($adCsObjectGroup)) -or (ReportTransient($aadCsObjectGroup)))
    {
        ReportOutput -PropertyName "Group - Transient Object" -PropertyValue "True"
    }	

    #
    # Check if group has synchronization issues due to "Maximum group member limit exceeded"
    #
    if ($adCsObjectGroup.SerializedXml -contains "Maximum Group member count exceeded" -or $aadCsObjectGroup.SerializedXml -contains "Maximum Group member count exceeded")
    {
        "Maximum group member count limit has been exceeded" | ReportError
        Write-Host "`r`n"
    }

    #
    # Get object graph for member
    #
    $adCsObjectMember = $null
    $mvObjectMember = $null
    $aadCsObjectMember = $null
    $objectGraphMember = CheckObjectGraph($GroupADConnectorName)($MemberDN)([ref]$adCsObjectMember)([ref]$mvObjectMember)([ref]$aadCsObjectMember)	

    #
    # Check attribute based filtering for member
    #
    CheckAttributeBasedFiltering($GroupADConnectorName)($MemberDN)($adCsObjectMember)($mvObjectMember)($aadCsObjectMember)
    
    # Check for Sync Errors and Export Errors on group member object
    if ((ReportSyncError($adCsObjectMember)) -or (ReportSyncError($aadCsObjectMember)))
    {
        ReportOutput -PropertyName "Sync Error(s) on group member" -PropertyValue "True"
    }

    if ((ReportExportError($adCsObjectMember)) -or (ReportExportError($aadCsObjectMember)))
    {
        ReportOutput -PropertyName "Export Error(s) on group member" -PropertyValue "True"
    }	

    # Check for transient on group member object
    if ((ReportTransient($adCsObjectMember)) -or (ReportTransient($aadCsObjectMember)))
    {
        ReportOutput -PropertyName "Group Member - Transient Object" -PropertyValue "True"
    }

    #
    # Validate member reference exists in AD CS
    #
    "Checking for group membership in sync engine..." | Write-Host -fore White
    $isMemberInAdCS = IsMemberOfGroupInCS($adCsObjectGroup)($adCsObjectMember)
    if (-not $isMemberInAdCS)
    {
        "Object `"$MemberDN`" is not a member of group `"$GroupDN`" in AD connector space" | ReportError
    }
    else
    {
        "Object `"$MemberDN`" is a member of group `"$GroupDN`" in AD connector space" | ReportOutput
    }

    #
    # Validate member reference exists in MV
    #
    $isMemberInMV = IsMemberOfGroupInMV($mvObjectGroup)($mvObjectMember)
    if (-not $isMemberInMV)
    {
        "Object `"$MemberDN`" is not a member of group `"$GroupDN`" in Metaverse" | ReportError
    }
    else
    {
        "Object `"$MemberDN`" is a member of group `"$GroupDN`" in Metaverse" | ReportOutput
    }

    #
    # Validate member reference exists in AAD CS
    #
    $isMemberInAadCS = IsMemberOfGroupInCS($aadCsObjectGroup)($aadCsObjectMember)
    if (-not $isMemberInAadCS)
    {
        "Object `"$MemberDN`" is not a member of group `"$GroupDN`" in AAD connector space" | ReportError
    }
    else
    {
        "Object `"$MemberDN`" is a member of group `"$GroupDN`" in AAD connector space" | ReportOutput
    }

    #
    # Set output directory for per-object html report
    #
    Set-OutputDirectory

    #
    # Per-object html report date-time
    #
    $reportDate = [string] $(Get-Date -Format yyyyMMddHHmmss)

    if ($SynchronizationIssueHtmlItems.Count -gt 0)
    {
        $synchronizationIssuesHtmlGroup = WriteHtmlAccordionGroup($SynchronizationIssueHtmlItems)($global:HtmlSynchronizationIssuesSectionTitle)
        $HtmlGroupList.Insert(0, $synchronizationIssuesHtmlGroup)
        $SynchronizationIssueHtmlItems.Clear()        
    }
    
    if ($HtmlGroupList.Count -gt 0)
    {
        $objectSyncDiagnosticsHtmlContent = WriteHtmlAccordion($HtmlGroupList)($GroupDN)

        $objectSyncDiagnosticsHtmlBody = WriteHtmlBody($objectSyncDiagnosticsHtmlContent)

        $objectSyncDiagnosticsHtml = WriteHtml($objectSyncDiagnosticsHtmlBody)

        Write-Host "`r`n"

        Export-ObjectDiagnosticsHtmlReport -Title $GroupDN -ReportDate $reportDate -HtmlDoc $objectSyncDiagnosticsHtml

        $HtmlGroupList.Clear()
    }

    Write-Host "`r`n"
    Write-Host "`r`n"

    #
    # Ask if the tool is helpful for the synchronization issues
    #
    # The customer is going to answer for each synchronization issue separately
    #
    AskIfToolHelpful($GroupDN)

    $SynchronizationIssueList.Clear()
}

Function ReportTransient
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$false)]
        [AllowNull()]
        $AdCSObject
    )

    if ($AdCSObject -ne $null -and $AdCSObject.IsTransient)
    {
        $AdCSObjectDN = $AdCSObject.DistinguishedName
        "Object `"$AdCSObjectDN`" is a transient object" | ReportWarning
        return $true
    }

    return $false
}

Function ReportSyncError
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$false)]
        [AllowNull()]
        $AdCSObject
    )

    if ($AdCSObject -ne $null -and $AdCSObject.HasSyncError)
    {
        $AdCSObjectDN = $AdCSObject.DistinguishedName
        "Object `"$AdCSObjectDN`" has synchronization error(s)" | ReportError

        if ($AdCsObject.ExportedChangedNotReimported)
        {
            "Object `"$AdCSObjectDN`" has exported-change-not-reimported error. This means that the changes exported to the connected data source is being overriden." | ReportError
        }

        return $true
    }

    return $false
}

Function ReportExportError
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$false)]
        [AllowNull()]
        $AdCSObject
    )

    if ($AdCSObject -ne $null -and $AdCSObject.HasExportError)
    {
        $AdCSObjectDN = $AdCSObject.DistinguishedName
        "Object `"$AdCSObjectDN`" has export error(s)" | ReportError
        return $true
    }

    return $false
}

Function GetObjectDN
{
    param
    (
        [string]
        [parameter(mandatory=$false)]
        $ObjectDN,

        [string]
        [parameter(mandatory=$false)]
        $PromptMessage
    )

    while ([string]::IsNullOrEmpty($ObjectDN))
    {
        $ObjectDN = Read-Host $PromptMessage
    }

    return $ObjectDN
}

Function CheckObjectGraph
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $ADConnectorName,

        [string]
        [parameter(mandatory=$true)]
        $ObjectDN,

        [parameter(mandatory=$false)]
        [ref]$AdCSObject,

        [parameter(mandatory=$false)]
        [ref]$MVObject,

        [parameter(mandatory=$false)]
        [ref]$AadCSObject
    )

    "Checking for object `"$ObjectDN`" in sync engine..." | Write-Host -fore White

    #
    # Get object in the AD connector space
    #
    $AdCSObject.Value = GetCSObject($ADConnectorName)($ObjectDN)

    if ($AdCSObject.Value -eq $null) 
    {
        "Object `"$ObjectDN`" is not found in AD Connector Space - `"$ADConnectorName`"" | ReportError

        return $false
    }
    else
    {
        "Object `"$ObjectDN`" is found in AD Connector Space - `"$ADConnectorName`"" | ReportOutput
    }

    #
    # Get object in the metaverse
    #
    if ($AdCSObject.Value -ne $null -and $AdCSObject.Value.ConnectedMVObjectId -ne [System.Guid]::Empty)
    {
        $MVObject.Value = GetMVObjectByIdentifier($AdCSObject.Value.ConnectedMVObjectId)
    }

    if ($MVObject.Value -eq $null)
    {
        "Object `"$ObjectDN`" is not found in Metaverse" | ReportError
    }
    else
    {
        "Object `"$ObjectDN`" is found in Metaverse" | ReportOutput
    }

    #
    # Get joined AD connector space objects
    #
    $joinedAdCsObjects = New-Object System.Collections.ArrayList
    if ($MVObject.Value -ne $null)
    {
        $links = $MVObject.Value.Lineage
        $joinedAdCsObjectLinks = $links.Where( {$_.ConnectedCsObjectId -ne $AdCSObject.Value.ObjectId -and $_.ConnectorId -ne "b891884f-051e-4a83-95af-2544101c9083"})
        $joinedAdCsObjectLinks | ForEach {
            $joinedAdCsObjectDN = $_.ConnectedCsObjectDN
            $joinedAdCsObjectConnectorName = $_.ConnectorName                
            $joinedAdCsObject = GetCSObjectByIdentifier($_.ConnectedCsObjectId)

            if ($joinedAdCsObject -eq $null)
            {
                "Could not find joined object `"$joinedAdCsObjectDN`" in the AD Connector Space `"$joinedAdCsObjectConnectorName`"" | ReportError
            }
            else
            {
                "Object `"$ObjectDN`" is joined to object `"$joinedAdCsObjectDN`" in the Metaverse." | ReportOutput
                $joinedAdCsObjects.Add($joinedAdCsObject)
            }
        }
    }

    #
    # Get joined AD objects
    #
    $joinedAdObjects = New-Object System.Collections.ArrayList
    if ($joinedAdCsObjects.Count -gt 0)
    {
        $joinedAdCsObjects | ForEach {
            $joinedAdCsObjectDN = $_.DistinguishedName
            $joinedAdCsObjectConnectorName = $_.ConnectorName    
            $joinedAdObject = Search-ADSyncDirectoryObjects -AdConnectorId $_.ConnectorId -LdapFilter "(distinguishedName=$joinedAdCsObjectDN)" -SearchScope Subtree -SizeLimit 1

            if ($joinedAdObject -eq $null -or $joinedAdObject.Count -lt 1)
            {
                "Could not find joined object `"$joinedAdCsObjectDN`" in on-premises AD `"$joinedAdCsObjectConnectorName`"." | ReportError
            }
            else
            {
                $joinedAdObjects.Add($joinedAdObject[0])
            }
        }
    }

    #
    # Get object in AAD connector space
    #
    if ($MVObject.Value -ne $null)
    {
        $aadConnector = GetAADConnector

        if ($aadConnector -ne $null)
        {
            $aadCsObjectId = GetTargetCSObjectId($MVObject.Value)($aadConnector.Identifier)
        
            if ($aadCsObjectId -ne $null)
            {
                $AadCSObject.Value = GetCSObjectByIdentifier($aadCsObjectId)
            }
        }
    }

    if ($AadCSObject.Value -eq $null)
    {
        "Object `"$ObjectDN`" is not found in AAD Connector Space" | ReportError
    }
    else
    {
        "Object `"$ObjectDN`" is found in AAD Connector Space" | ReportOutput
    }

    Write-Host "`r`n"

    if ($AdCsObject.Value -ne $null -and $MVObject.Value -ne $null -and $AadCsObject.Value -ne $null)
    {
        return $true
    }

    return $false
}

Function GetAADDomains
{
    if ($isNonInteractiveMode)
    {
        return $null
    }

    return Get-MsolDomain
}

Function ConnectToAAD
{
    if ($isNonInteractiveMode)
    {
        return
    }

    #
    # Get AAD Tenant credentials if not specified as an input parameter
    #
    if ($global:AADTenantCredential -eq $null)
    {
        $global:AADTenantCredential = Get-Credential -Message "Please enter Azure AD Tenant global administrator or hybrid identity administrator credentials:"    
    }

    Write-Host "`r`n"

    #
    # Connect to AAD Tenant
    #
    try
    {
        "Connecting to Azure AD Tenant..." | Write-Host -fore White

        Connect-MsolService -Credential $global:AADTenantCredential -ErrorAction Stop

        "OK" | Write-Host -fore Green
        Write-Host "`r`n"
    }
    catch
    {
        $global:AADTenantCredential = $null

        "Failed to connect to AAD Tenant. Details: $($_.Exception.Message)" | Write-Host -fore Red
        Write-Host "`r`n"
        return
    }
}


Function CheckDomainBasedFiltering
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.ConnectorPartition]
        [parameter(mandatory=$false)]
        $Partition,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector]
        [parameter(mandatory=$true)]
        $AdConnector,

        [string]
        [parameter(mandatory=$true)]
        $ObjectDN
    )

    if ($Partition -eq $null)
    {
        return
    }

    $htmlMessageList = New-Object System.Collections.Generic.List[string]

    "Checking Domain Filtering configuration..." | Write-Host

    $partitionDN = $Partition.DN
    $adConnectorName = $AdConnector.Name
    
    if (-Not $Partition.Selected)
    {
        WriteEventLog($EventIdDomainFiltered)($EventMsgDomainFiltered -f ($ObjectDN, $partitionDN))

        $SynchronizationIssueList.Add($global:DomainFilteringIssue)

        Write-Host "`r`n"
        "DOMAIN FILTERING - ANALYSIS:" | Write-Host -fore Cyan
        "----------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        $consoleMessage = "Object is not present in sync scope. Object belongs to domain $partitionDN which is filtered from synchronization." 
        $consoleMessage | ReportError -PropertyName "Domain Filtered" -PropertyValue "True"
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        "DOMAIN FILTERING - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
        "---------------------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $message = "Include the partition $partitionDN in the list of domains that should be synced. To read more on how to do this, please see:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:DomainBasedFilteringUrl) 
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:DomainBasedFilteringUrl)($global:DomainBasedFilteringText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 0
        $htmlMessageList.Add($htmlMessage)


        $domainFilteringHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("Domain Filtering")

        $SynchronizationIssueHtmlItems.Add($domainFilteringHtmlItem)

        return
    }

    $runProfiles = $AdConnector.RunProfiles
    
    if ($runProfiles -eq $null)
    {
        WriteEventLog($EventIdDomainFiltered)($EventMsgDomainFiltered -f ($ObjectDN, $partitionDN))

        $SynchronizationIssueList.Add($global:DomainFilteringIssue)

        Write-Host "`r`n"
        "DOMAIN FILTERING - ANALYSIS:" | Write-Host -fore Cyan
        "----------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        $consoleMessage = "No run profiles are configured on the AD Connector $adConnectorName." 
        $consoleMessage | ReportError -PropertyName "Domain Filtered" -PropertyValue "True"
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        "DOMAIN FILTERING - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
        "---------------------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $message = "Update the run profiles to include the partition $partitionDN that should be synced. To read more on how to do this, please see:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:DomainBasedFilteringUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:DomainBasedFilteringUrl)($global:DomainBasedFilteringText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 0
        $htmlMessageList.Add($htmlMessage)


        $domainFilteringHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("Domain Filtering")

        $SynchronizationIssueHtmlItems.Add($domainFilteringHtmlItem)

        return
    }

    $runProfilesMissingStep = @()
    foreach ($runProfile in $runProfiles)
    {
        $runSteps = $runProfile.RunSteps
        if ($runSteps -eq $null)
        {
            $runProfilesMissingStep += $runProfile.Name
            continue
        }

        $foundRunStep = $false
        foreach ($runStep in $runSteps)
        {
            if ($runStep.PartitionIdentifier -eq $Partition.Identifier)
            {
                $foundRunStep = $true
                break
            }
        }

        if (-Not $foundRunStep)
        {
            $runProfilesMissingStep += $runProfile.Name		
        }
    }

    if ($runProfilesMissingStep.Count -gt 0)
    {
        WriteEventLog($EventIdDomainFiltered)($EventMsgDomainFiltered -f ($ObjectDN, $partitionDN))

        $SynchronizationIssueList.Add($global:DomainFilteringIssue)

        $runProfilesMissingStepOutput = $runProfilesMissingStep -join ","
        
        Write-Host "`r`n"

        "DOMAIN FILTERING - ANALYSIS:" | Write-Host -fore Cyan
        "----------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        $consoleMessage = "No run steps are configured for directory partition $partitionDN for some run profile(s), namely: $runProfilesMissingStepOutput."
        $consoleMessage | ReportError -PropertyName "Domain Filtered" -PropertyValue "True"
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        "DOMAIN FILTERING - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
        "---------------------------------------" | Write-Host -fore Cyan
        
        $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)
        

        $message = "Add steps to the above run profile(s) to include the partition $partitionDN. To read more on how to do this, please see:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:DomainBasedFilteringUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:DomainBasedFilteringUrl)($global:DomainBasedFilteringText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 0
        $htmlMessageList.Add($htmlMessage)


        $domainFilteringHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("Domain Filtering")

        $SynchronizationIssueHtmlItems.Add($domainFilteringHtmlItem)
    }
    else
    {
        "There is no domain level filtering configuration that prevents this object from being imported in the AD connector space." | Write-Host
        Write-Host "`r`n"
        ReportOutput -PropertyName "Domain Filtered" -PropertyValue "False"
    }
}

Function Get-DirectoryPartitionDN
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty]
        $ObjectDN
    )

    $index = $ObjectDN.IndexOf("DC=")
    return $ObjectDN.Substring($index)
}

Function DiagnoseGroupObject
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector]
        [parameter(mandatory=$true)]
        $AdConnector,

        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdCsObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $MvObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadCsObject,

        [Microsoft.Online.Administration.Group]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadGroupObject
    )

    WriteEventLog($EventIdDiagnoseSingleObject)($EventMsgDiagnoseSingleObject)

    "Checking group member count limitation..." | Write-Verbose

    # check member count
    if ($AdCsObject -ne $null -and $AdCsObject.HasSyncError -and $AdCsObject.SerializedXml -contains "Maximum Group member count exceeded")
    {
        $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        $consoleMessage = "For the given group object the maximum member count limit has been exceeded. Member attribute is not going to be synchronized to Azure Active Directory."
        $consoleMessage | ReportError -PropertyName "Maximum member count limit exceeded" -PropertyValue "True"
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)

        $groupMemberCountHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("Group Member Count")
        $SynchronizationIssueHtmlItems.Add($groupMemberCountHtmlItem)
    }
    else
    {
        "OK" | Write-Verbose
    }
}

Function DiagnoseFSP
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector]
        [parameter(mandatory=$true)]
        $AdConnector,

        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdCsObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $MvObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadCsObject
    )

    WriteEventLog($EventIdDiagnoseSingleObject)($EventMsgDiagnoseSingleObject)

    if ($AdCsObject -ne $null -and $MvObject -eq $null)
    {
        $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        $consoleMessage = "FSP is unable to find an exisitng metaverse object to join to. This typically happens if the primary object is not synchronized."
        $consoleMessage | ReportError
        Write-Host "`r`n"
    }
}

Function DiagnoseUserObject
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector]
        [parameter(mandatory=$true)]
        $AdConnector,

        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdCsObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $MvObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadCsObject,

        [Microsoft.Online.Administration.User]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadUserObject,

        [string]
        [parameter(mandatory=$true)]
        [AllowNull()]
        [AllowEmptyString()]
        $AadTenantDefaultDomainName, 

        [Microsoft.Online.Administration.Domain[]]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadTenantDomains
    )

    WriteEventLog($EventIdDiagnoseSingleObject)($EventMsgDiagnoseSingleObject)

    #
    # Diagnose linked mailbox issues
    #
    if ($AdCsObject -and $MvObject)
    {
        DiagnoseUserObjectLinkedMailBox($AdCsObject)($MvObject)
    }

    if ($AadUserObject -eq $null)
    {
        return
    }

    "Checking UserPrincipalName (UPN) Mismatch..." | Write-Host -fore White

    #
    # Get "userPrincipalName" attribute value of the metaverse object
    #
    $userPrincipalNameAttribute = GetMVObjectAttribute($MvObject)("userPrincipalName")
    $userPrincipalNameValue = $null

    if ($userPrincipalNameAttribute -ne $null)
    {
        $userPrincipalNameValue = $userPrincipalNameAttribute.Values[0]
    }
    else
    {
        "Failed to get `"userPrincipalName`" attribute value of AADConnect object." | Write-Host -fore Red

        Write-Host "`r`n"
        return
    }

    #
    # Check if there is a mismatch between MV object UPN and AAD Tenant object UPN
    #
    if ($userPrincipalNameValue -ne $AadUserObject.UserPrincipalName)
    {
        Write-Host "`r`n"

        "USERPRINCIPALNAME MISMATCH - ANALYSIS:" | Write-Host -fore Cyan
        "--------------------------------------" | Write-Host -fore Cyan

        "For the given user object there is a `"userPrincipalName`" attribute value mismatch between AADConnect and AAD Tenant." | Write-Host -fore Yellow
        Write-Host "`r`n"

        DiagnoseUPNMismatch($MvObject)($AadUserObject)($AadTenantDefaultDomainName)($AadTenantDomains)
    }
    else
    {
        "OK" | Write-Host -fore Green
        Write-Host "`r`n"
    }
}

#
# Provisioning a user object having exchange recipient type as "Linked Mailbox" results in a metaverse object with NULL source anchor.
# By default, a metaverse object with NULL source anchor does NOT flow into AAD connector space.
#
# The exchange recipient type "Linked Mailbox" is specified with the on-premises AD attribute "msExchRecipientTypeDetails" having value 2.
#
# In order not to end up with a metaverse object having NULL source anchor:
#
#    1- There should be an active account (not disabled, userAccountControl != 2) in another account forest which joins into the same metaverse object.
#       So that source anchor would flow from the active account.
#
#    2- If there is no active account joining into the same metaverse object, then the customer should consider converting the "Linked Mailbox" to a "User Mailbox".
#       The exchange recipient type "User Mailbox" is specified with the on-premises AD attribute "msExchRecipientTypeDetails" having value 1.
#
Function DiagnoseUserObjectLinkedMailBox
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        $AdCsObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        $MvObject
    )

    $htmlMessageList = New-Object System.Collections.Generic.List[string]

    #
    # Get AD connector space object attribute "msExchRecipientTypeDetails".
    #
    $msExchRecipientTypeDetailsAttribute = GetCSObjectAttribute($AdCsObject)("msExchRecipientTypeDetails")
    $msExchRecipientTypeDetailsValue = $null

    if ($msExchRecipientTypeDetailsAttribute -ne $null)
    {
        $msExchRecipientTypeDetailsValue = $msExchRecipientTypeDetailsAttribute.Values[0]
    }
    else
    {
        #
        # Attribute "msExchRecipientTypeDetails" is not imported from on-premises AD.
        #

        return
    }
    
    #
    # Exchage recipient type is NOT "Linked Mailbox"
    #
    $msExchRecipientTypeDetailsValueInt=0
    if (([System.Int32]::TryParse($msExchRecipientTypeDetailsValue, [ref] $msExchRecipientTypeDetailsValueInt)) -or ($msExchRecipientTypeDetailsValueInt -ne 2))
    {
        return
    }

    "Checking Linked Mailbox related issues..." | Write-Host -fore White

    #
    # Get metaverse object attribute "sourceAnchor"
    #
    $sourceAnchorAttribute = GetMVObjectAttribute($MvObject)("sourceAnchor")

    #
    # The metaverse object is NOT going to flow into AAD connector space.
    #
    if ($sourceAnchorAttribute -eq $null)
    {
        WriteEventLog($EventIdLinkedMailboxIssue)($EventMsgLinkedMailboxIssue)

        $SynchronizationIssueList.Add($global:LinkedMailboxIssue)

        Write-Host "`r`n"

        "LINKED MAILBOX ISSUE - ANALYSIS:" | Write-Host -fore Cyan
        "--------------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        ReportOutput -PropertyName "Linked Mailbox Issue" -PropertyValue "True"
        $consoleMessage = "Azure AD Connect cannot synchronize the given user object to Azure AD tenant since the object has on-premises linked mailbox."
        $consoleMessage | ReportError
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "The exchange recipient type `"Linked Mailbox`" is specified with the on-premises AD attribute `"msExchRecipientTypeDetails`" having value 2." 
        $consoleMessage | Write-Host -fore Red
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        "LINKED MAILBOX ISSUE - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
        "-------------------------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "1- Convert the linked mailbox to a user mailbox." 
        $consoleMessage | Write-Host -fore White
        
        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "In case the given user object is represented only once across all on-premises directories, then you may consider converting the linked mailbox to a user mailbox to unblock the synchronization."
        $consoleMessage | Write-Host -fore Yellow

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)
        

        $message = "In order to learn how to convert the mailbox, please see:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:ConvertLinkedMailboxUrl)
        $consoleMessage | Write-Host -fore Yellow

        $htmlMessage = GetHtmlMessageWithLink($message)($global:ConvertLinkedMailboxUrl)($global:ConvertLinkedMailboxText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        
        $message = "In order to learn how objects are represented only once across all on-premises directories, please see `"Uniquely identifying your users`" section at:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:InstallationUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:InstallationUrl)($global:InstallationText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "2- Create an active user account in another on-premises account forest joining into the same user object." 
        $consoleMessage | Write-Host -fore White

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)
        

        $consoleMessage = "A linked mailbox is supposed to be associated with a disabled user account or contact object joining into an active user account from another on-premises AD forest."
        $consoleMessage | Write-Host -fore Yellow
        
        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        
        $message = "In order to learn how objects join from different on-premises forests, please see `"Uniquely identifying your users`" section at:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:InstallationUrl)
        $consoleMessage | Write-Host -fore Yellow
        
        $htmlMessage = GetHtmlMessageWithLink($message)($global:InstallationUrl)($global:InstallationText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)
        
        
        $message = "In order to learn about account-resource forest topology, please see `"Multiple forests: account-resource forest`" section at:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:TopologiesUrl)
        $consoleMessage | Write-Host -fore Yellow

        $htmlMessage = GetHtmlMessageWithLink($message)($global:TopologiesUrl)($global:TopologiesText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $linkedMailboxHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("Linked Mailbox")

        $SynchronizationIssueHtmlItems.Add($linkedMailboxHtmlItem)
    }
    else
    {
        ReportOutput -PropertyName "Linked Mailbox Issue" -PropertyValue "False"
        "OK" | Write-Host -fore Green
        Write-Host "`r`n"
    }
}

Function DiagnoseUPNMismatch
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        $MvObject,

        [Microsoft.Online.Administration.User]
        [parameter (mandatory=$true)]
        $AadTenantObject,

        [string]
        [parameter (mandatory=$true)]
        $AadTenantDefaultDomainName,

        [Microsoft.Online.Administration.Domain[]]
        [parameter(mandatory=$true)]
        $AadTenantDomains
    )

    $htmlMessageList = New-Object System.Collections.Generic.List[string]

    #
    # Report metaverse object UPN
    #
    $mvObjectUserPrincipalNameAttribute = GetMVObjectAttribute($MvObject)("userPrincipalName")
    $mvObjectUserPrincipalNameValue = $null

    if ($mvObjectUserPrincipalNameAttribute -ne $null)
    {
        $mvObjectUserPrincipalNameValue = $mvObjectUserPrincipalNameAttribute.Values[0]
    }
    else
    {
        "Failed to get `"userPrincipalName`" attribute value of the metaverse object." | Write-Host -fore Red

        Write-Host "`r`n"
        return
    }

    $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)

    
    $consoleMessage = "For the given user object there is a `"userPrincipalName`" attribute value mismatch between AADConnect and AAD Tenant."

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "AADConnect object `"userPrincipalName`" attribute value is:"
    $consoleMessage | Write-Host -fore Cyan

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = $mvObjectUserPrincipalNameValue
    $consoleMessage | Write-Host -fore Green
    Write-Host "`r`n"

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    #
    # Report AAD Tenant object UPN
    #
    $aadTenantObjectUserPrincipalNameValue = $AadTenantObject.UserPrincipalName


    $consoleMessage = "AAD Tenant object `"userPrincipalName`" attribute value is:"
    $consoleMessage | Write-Host -fore Cyan

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = $aadTenantObjectUserPrincipalNameValue
    $consoleMessage | Write-Host -fore Green

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    WriteEventLog($EventIdUPNMismatch)($EventMsgUPNMismatch -f ($mvObjectUserPrincipalNameValue, $aadTenantObjectUserPrincipalNameValue))

    Write-Host "`r`n"

    $mvObjectUpnPrefix = $mvObjectUserPrincipalNameValue.Split('@')[0]
    $mvObjectUpnSuffix = $mvObjectUserPrincipalNameValue.Split('@')[1]

    $aadTenantObjectUpnSuffix = $aadTenantObjectUserPrincipalNameValue.Split('@')[1]

    #
    # Get AAD Tenant object domain
    #
    $aadTenantObjectDomain = $AadTenantDomains | Where {$_.Name -eq $aadTenantObjectUpnSuffix}

    #
    # Azure AD does NOT allow updates to UserPrincipalName in case following conditions occur:
    #
    #     + AAD Tenant user object is MANAGED (not federated)
    #     + AAD Tenant user object is LICENSED               
    #     + AAD Tenant "SynchronizeUpnForManagedUsers" feature is DISABLED
    #
    #
    # LIMITATION: Currently there is NO way to determine authentication type for the AAD Tenant user object. Therefore, checking the authenticaton type (managed/federated)
    #             of the domain that the AAD Tenant user object is part of until finding a way to specifically detecting the authentication type for the AAD Tenant user object.
    #
    #             As an example, today it is possible to have an AAD Tenant domain with managed authentication having federated user accounts.
    #
    $isSynchronizeUpnForManagedUserFeatureEnabled = IsSynchronizeUpnForManagedUsersFeatureEnabled

    $isAadTenantObjectUPNUpdatesBlocked = (($aadTenantObjectDomain.Authentication -eq "Managed") -and $AadTenantObject.IsLicensed -and !$isSynchronizeUpnForManagedUserFeatureEnabled)

    #
    # Get on-premises attribute name configured as source of metaverse object UPN
    #
    $upnAttributeName = (Get-ADSyncGlobalSettings).Parameters["Microsoft.SynchronizationOption.UPNAttribute"].Value

    if ($upnAttributeName -ne "userPrincipalName")
    {
        "Alternate Login ID:" | Write-Host -fore Cyan
        "On-premises attribute `"$($upnAttributeName)`" is configured as source of Azure AD username." | Write-Host -fore Green
        Write-Host "`r`n"
    }
    
    #
    # Report that UPN update is blocked for the AAD Tenant user object.
    #
    if ($isAadTenantObjectUPNUpdatesBlocked)
    {
        WriteEventLog($EventIdUPNUpdateBlocked)($EventMsgUPNUpdateBlocked -f $aadTenantObjectUserPrincipalNameValue)

        $SynchronizationIssueList.Add($global:UPNMismatchDirSyncFeatureIssue)

        $consoleMessage = "Updates to `"userPrincipalName`" attribute of the AAD Tenant object `"$($aadTenantObjectUserPrincipalNameValue)`" is blocked since AAD Tenant DirSync feature `"SynchronizeUpnForManagedUsers`" is disabled."
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        "USERPRINCIPALNAME MISMATCH - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
        "-------------------------------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "AAD Tenant DirSync feature `"SynchronizeUpnForManagedUsers`" should be enabled to unblock updates."
        $consoleMessage | Write-Host -fore Yellow
        
        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "Once AAD Tenant DirSync feature `"SynchronizeUpnForManagedUsers`" is enabled, it cannot be disabled."
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        $message = "For more information, please see `"Scenario 2`" in the following:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:UPNMismatchUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:UPNMismatchUrl)($global:UPNMismatchText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        $upnMismatchHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("UserPrincipalName Mismatch")

        $SynchronizationIssueHtmlItems.Add($upnMismatchHtmlItem)

        return
    }

    #
    # Check if metaverse object UPN suffix is verified as an internet-routable domain name in AAD Tenant.
    #
    $isMvObjectUpnSuffixVerified = IsUPNSuffixVerifiedInAADTenant($MvObjectUpnSuffix)($AadTenantDomains)

    if ($isMvObjectUpnSuffixVerified)
    {
        WriteEventLog($EventIdUPNSuffixVerified)($EventMsgUPNSuffixVerified -f ($mvObjectUpnSuffix, $AadTenantDefaultDomainName))

        $updatedAadTenantDomain = Get-MsolDomain -DomainName $mvObjectUpnSuffix

        if ($updatedAadTenantDomain.Authentication -eq "Federated" -and $aadTenantObjectDomain.Authentication -eq "Federated")
        {
            WriteEventLog($EventIdUPNFederatedDomainChange)($EventMsgUPNFederatedDomainChange -f ($mvObjectUserPrincipalNameValue, $aadTenantObjectUserPrincipalNameValue))

            $SynchronizationIssueList.Add($global:UPNMismatchFederatedDomainChangeIssue)

            $consoleMessage = "Azure Active Directory does NOT allow to synchronize UPN suffix change from one federated domain to another federated domain."
            $consoleMessage | Write-Host -fore Yellow
            Write-Host "`r`n"

            $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
            $htmlMessageList.Add($htmlMessage)


            "USERPRINCIPALNAME MISMATCH - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
            "-------------------------------------------------" | Write-Host -fore Cyan
            
            $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
            $htmlMessageList.Add($htmlMessage)

            
            $message = "To change from one federated domain to another, please see:"
            $consoleMessage = GetConsoleMessageWithLink($message)($global:FederatedDomainChangeUrl)
            $consoleMessage | Write-Host -fore Yellow
            Write-Host "`r`n"

            $htmlMessage = GetHtmlMessageWithLink($message)($global:FederatedDomainChangeUrl)($global:FederatedDomainChangeText)
            $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
            $htmlMessageList.Add($htmlMessage)
        }
    }
    else
    {
        WriteEventLog($EventIdUPNSuffixNotVerified)($EventMsgUPNSuffixNotVerified -f ($mvObjectUpnSuffix, $AadTenantDefaultDomainName))

        $SynchronizationIssueList.Add($global:UPNMismatchNonVerifiedUpnSuffixIssue)

        $consoleMessage = "UPN suffix `"$($mvObjectUpnSuffix)`" is NOT verified with AAD Tenant `"$($AadTenantDefaultDomainName)`" as a domain."
        $consoleMessage | Write-Host -fore Yellow
        
        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        
        $consoleMessage = "Azure Active Directory replaces such UPN suffixes with default domain name `"onmicrosoft.com`"."
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "When the UPN suffix is NOT verified with the AAD Tenant, Azure AD takes the following inputs into account in the given order to calculate the UPN prefix in the cloud:"
        $consoleMessage | Write-Host -fore Yellow
        
        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        
        "`t1- On-Premises `"mailNickName`" attribute" | Write-Host -fore Yellow
        "`t2- Primary SMTP Address" | Write-Host -fore Yellow
        "`t3- On-Premises `"mail`" attribute" | Write-Host -fore Yellow
        "`t4- UserPrincipalName/Alternate Login ID" | Write-Host -fore Yellow
        Write-Host "`r`n"


        $htmlMessage = WriteHtmlMessage -message "1- On-Premises `"mailNickName`" attribute" -color "#252525" -paddingLeft "30px" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        $htmlMessage = WriteHtmlMessage -message "2- Primary SMTP Address" -color "#252525" -paddingLeft "30px" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)
     
        $htmlMessage = WriteHtmlMessage -message "3- On-Premises `"mail`" attribute" -color "#252525" -paddingLeft "30px" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)
        
        $htmlMessage = WriteHtmlMessage -message "4- UserPrincipalName/Alternate Login ID" -color "#252525" -paddingLeft "30px" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)
            
         
        $message = "Please see to understand how Azure Active Directory populates UPN in the cloud:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:CloudUPNPopulationUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:CloudUPNPopulationUrl)($global:CloudUPNPopulationText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        "USERPRINCIPALNAME MISMATCH - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
        "-------------------------------------------------" | Write-Host -fore Cyan
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "Please consider to follow one of the options given below:"
        $consoleMessage | Write-Host -fore White
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "1- Verify current UPN suffix with AAD Tenant"
        $consoleMessage | Write-Host -fore White

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)
        
        
        $consoleMessage = "In order to use `"$($mvObjectUserPrincipalNameValue)`" as Azure Active Directory username, you need to verify UPN suffix `"$($mvObjectUpnSuffix)`" with AAD Tenant `"$($AadTenantDefaultDomainName)`"."
        $consoleMessage | Write-Host -fore Yellow

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $message = "In order to verify a domain name with AAD Tenant, please see:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:VerifyDomainNameUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:VerifyDomainNameUrl)($global:VerifyDomainNameText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        if ($upnAttributeName -eq "userPrincipalName")
        {
            $consoleMessage = "2- Alternative UPN Suffix"
            $consoleMessage | Write-Host -fore White

            $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
            $htmlMessageList.Add($htmlMessage)


            $consoleMessage = "As another option, you may consider adding an alternative UPN suffix to your on-premises accounts."
            $consoleMessage | Write-Host -fore Yellow 

            $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
            $htmlMessageList.Add($htmlMessage)

            
            $message = "Please see:"
            $consoleMessage = GetConsoleMessageWithLink($message)($global:AlternativeUPNSuffixUrl)
            $consoleMessage | Write-Host -fore Yellow
            Write-Host "`r`n"

            $htmlMessage = GetHtmlMessageWithLink($message)($global:AlternativeUPNSuffixUrl)($global:AlternativeUPNSuffixText)
            $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
            $htmlMessageList.Add($htmlMessage)


            $consoleMessage = "3- Run `"Set-MsolUserPrincipalName`" Cmdlet"
            $consoleMessage | Write-Host -fore White

            $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
            $htmlMessageList.Add($htmlMessage)
        }
        else
        {
            $consoleMessage = "2- Run `"Set-MsolUserPrincipalName`" Cmdlet"
            $consoleMessage | Write-Host -fore White

            $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
            $htmlMessageList.Add($htmlMessage)
        }


        $consoleMessage = "Apart from the UPN mismatch, if you are only interested in changing AAD Tenant user object `"userPrincipalName`" attribute prefix and/or suffix, then run AAD PowerShell cmdlet `"Set-MsolUserPrincipalName`"."
        $consoleMessage | Write-Host -fore Yellow

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $message = "In order to learn about the cmdlet, please see:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:SetUPNCmdletUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:SetUPNCmdletUrl)($global:SetUPNCmdletText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)
    }

    $upnMismatchHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("UserPrincipalName Mismatch")

    $SynchronizationIssueHtmlItems.Add($upnMismatchHtmlItem)
}

Function DiagnoseUserObjectProxyAddresses
{
    $htmlMessageList = New-Object System.Collections.Generic.List[string]

    $SynchronizationIssueList.Add($global:ChangePrimaryEmailAddress)

    "CHANGING EXCHANGE ONLINE PRIMARY EMAIL ADDRESS - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
    "---------------------------------------------------------------------" | Write-Host -fore Cyan

    $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)

    
    $consoleMessage = "1- Open `"ADSI Edit`" within a domain controller storing the target object in your on-premises directory."
    $consoleMessage | Write-Host -fore White
    Write-Host "`r`n"

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "2- Locate the target object."
    $consoleMessage | Write-Host -fore White
    Write-Host "`r`n"

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "3- Right click on the target object and go to `"Properties`"."
    $consoleMessage | Write-Host -fore White
    Write-Host "`r`n"

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "4- Find `"proxyAddresses`" attribute in the Attribute Editor."
    $consoleMessage | Write-Host -fore White
    Write-Host "`r`n"

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "5- Add/Change primary smtp address."
    $consoleMessage | Write-Host -fore White

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "Primary smtp address should have the prefix `"SMTP:`" in capital letters. Example: `"SMTP:myaccount@testdomain.com`"."
    $consoleMessage | Write-Host -fore Yellow

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "The domain part of the email address should be a domain name verified with the Azure AD tenant and it should NOT be the default tenant domain name ending with `"onmicrosoft.com`"."
    $consoleMessage | Write-Host -fore Yellow
    Write-Host "`r`n"

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "6- Find `"mail`" attribute in the Attribute Editor."
    $consoleMessage | Write-Host -fore White
    Write-Host "`r`n"

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "7- Set/Change `"mail`" attribute value to the same email address WITHOUT the prefix `"SMTP:`"."
    $consoleMessage | Write-Host -fore White
    Write-Host "`r`n"

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "8- At the end of next sync cycle, the primary email address in the cloud will change to the primary smtp address."
    $consoleMessage | Write-Host -fore White
    Write-Host "`r`n"

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "IMPORTANT:" 
    $consoleMessage | Write-Host -fore Yellow

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "Azure Active Directory does NOT allow to synchronize email addresses from the on-premises directory in case:"
    $consoleMessage | Write-Host -fore Yellow

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "a- the domain part of the on-premises email address is NOT a domain name verified with the Azure AD tenant."
    $consoleMessage | Write-Host -fore Yellow

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)


    $consoleMessage = "b- the domain part of the on-premises email address is the default tenant domain name ending with `"onmicrosoft.com`"."
    $consoleMessage | Write-Host -fore Yellow

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)


    $primaryEmailAddressHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)($global:ChangePrimaryEmailAddress)

    $SynchronizationIssueHtmlItems.Add($primaryEmailAddressHtmlItem)

    Write-Host "`r`n"
    Write-Host "`r`n"
}

Function DiagnoseUserObjectHideFromAddressLists
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector]
        [parameter(mandatory=$true)]
        $AdConnector,

        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        $AdObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        $AdCsObject
    )

    $htmlMessageList = New-Object System.Collections.Generic.List[string]

    $SynchronizationIssueList.Add($global:HideFromGlobalAddressList)

    "HIDING MAILBOX FROM EXCHANGE ONLINE GLOBAL ADDRESS LIST  - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
    "-------------------------------------------------------------------------------" | Write-Host -fore Cyan

    $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)


    #
    # Get AD connector space object attribute "msExchHideFromAddressLists"
    #
    $msExchHideFromAddressListsAttribute = GetCSObjectAttribute($AdCsObject)("msExchHideFromAddressLists")
    $msExchHideFromAddressListsValue = $null
    
    if ($msExchHideFromAddressListsAttribute -ne $null)
    {
        $msExchHideFromAddressListsValue = $msExchHideFromAddressListsAttribute.Values[0]
    }

    #
    # Get AD connector space object attribute "mailNickname"
    #
    $mailNickNameAttribute = GetCSObjectAttribute($AdCsObject)("mailNickname")
    $mailNickNameValue = $null

    if ($mailNickNameAttribute -ne $null)
    {
        $mailNickNameValue = $mailNickNameAttribute.Values[0]
    }

    $step = 1;

    #
    # Synchronizing "msExchHideFromAddressLists" attribute requires "mailNickname" attribute to be populated within on-premises directory.
    # 
    if ($msExchHideFromAddressListsValue -ne $null -and [bool] $msExchHideFromAddressListsValue -eq $true -and $mailNickNameValue -ne $null)
    {
        $consoleMessage = "$($step)- Your settings are correct. Please open a service request through Azure Portal or Office 365 Portal."
        $consoleMessage | Write-Host -fore White
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        $message = "`tMicrosoft Azure Portal:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:AzurePortalSupportBladeUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = WriteHyperlink($global:AzurePortalSupportBladeUrl)($global:AzurePortalSupportBladeText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        $message = "`tOffice 365 Portal:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:OfficePortalUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = WriteHyperlink($global:OfficePortalUrl)($global:OfficePortalText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        $hideFromAddressListHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("Hiding Account From Global Address List")

        $SynchronizationIssueHtmlItems.Add($hideFromAddressListHtmlItem)

        return
    }

    #
    # In order to hide account from exchange online global address list, "msExchHideFromAdressLists" attribute value must be TRUE.
    #
    if ($msExchHideFromAddressListsValue -ne $null -and [bool] $msExchHideFromAddressListsValue -eq $false)
    {
        $consoleMessage = "$($step)- Please set `"msExchHideFromAddressLists`" attribute value to `"TRUE`" within your on-premises directory."
        $consoleMessage | Write-Host -fore White
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)

        $step++
    }
    elseif ($msExchHideFromAddressListsValue -eq $null)
    {
        #
        # Check if "msExchHideFromAddressLists" attribute is included in the schema of the given AD Connector.
        #
        $adConnectorSchemaAttributeType = $AdConnector.Schema.AttributeTypes["msExchHideFromAddressLists"]

        #
        # Get on-premises AD object "msExchHideFromAddressLists" attribute value.
        #
        $adObjectAttributeValue = $null

        if ($AdObject.ContainsKey("msexchhidefromaddresslists"))
        {
            $adObjectAttributeValue = [bool] ([System.Collections.ArrayList] $AdObject["msexchhidefromaddresslists"])[0]
        }

        #
        # Recommended steps based on availability of "msExchHideFromAddressLists" attribute in AD Connector schema.
        #
        if ($adConnectorSchemaAttributeType -eq $null)
        {
            if ($adObjectAttributeValue -eq $null)
            {
                #
                # Customer needs to set the "msExchHideFromAddressLists" attribute value in the on-premises directory. 
                #
                # If the Active Directory schema does NOT have the Exchange attributes, then the customer needs to extend the schema. 
                #
                $consoleMessage = "$($step)- Please set `"msExchHideFromAddressLists`" attribute value to `"TRUE`" within your on-premises directory."
                $consoleMessage | Write-Host -fore White

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
                $htmlMessageList.Add($htmlMessage)


                $consoleMessage = "If the object does NOT have the attribute, then you need to extend your on-premises Active Directory schema to include the Exchange attributes."
                $consoleMessage | Write-Host -fore Yellow

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
                $htmlMessageList.Add($htmlMessage)


                $message = "In order to extend Active Directory schema, please see:"
                $consoleMessage = GetConsoleMessageWithLink($message)($global:ExtendSchemaUrl)
                $consoleMessage | Write-Host -fore Yellow
                Write-Host "`r`n"

                $htmlMessage = GetHtmlMessageWithLink($message)($global:ExtendSchemaUrl)($global:ExtendSchemaText)
                $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
                $htmlMessageList.Add($htmlMessage)

                $step++
            }
            elseif ($adObjectAttributeValue -eq $false)
            {
                $consoleMessage = "$($step)- Please set `"msExchHideFromAddressLists`" attribute value to `"TRUE`" within your on-premises directory."
                $consoleMessage | Write-Host -fore White
                Write-Host "`r`n"

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
                $htmlMessageList.Add($htmlMessage)

                $step++
            }
            
            #
            # As the AD Connector schema does NOT have the "msExchHideFromAddressLists" attribute, the customer needs to refresh the schema.
            #
            $consoleMessage = "$($step)- Please refresh schema stored in AADConnect for the on-premises directory `"$($AdConnector.Name)`"."
            $consoleMessage | Write-Host -fore White

            $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
            $htmlMessageList.Add($htmlMessage)


            $message = "In order to refresh schema from AADConnect Wizard, please see:"
            $consoleMessage = GetConsoleMessageWithLink($message)($global:RefreshSchemaUrl)
            $consoleMessage | Write-Host -fore Yellow
            Write-Host "`r`n"

            $htmlMessage = GetHtmlMessageWithLink($message)($global:RefreshSchemaUrl)($global:RefreshSchemaText)
            $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
            $htmlMessageList.Add($htmlMessage)

            $step++
        }
        else # "msExchHideFromAddressLists" attribute is available in the AD Connector schema.
        {
            #
            # Check if "msExchHideFromAddressLists" attribute is added to attribute inclusion list of the AD Connector.
            #
            $inclusionListAttribute = $AdConnector.AttributeInclusionList | ? {$_ -eq "msExchHideFromAddressLists"}

            if ($inclusionListAttribute -eq $null)
            {
                $consoleMessage = "$($step)- Please add `"msExchHideFromAddressLists`" attribute into attribute inclusion list of the AD Connector. Follow these steps:"
                $consoleMessage | Write-Host -fore White

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 1
                $htmlMessageList.Add($htmlMessage)


                $consoleMessage = "`ta- Open `"Synchronization Service Manager`" UI."
                $consoleMessage | Write-Host -fore Yellow

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
                $htmlMessageList.Add($htmlMessage)


                $consoleMessage = "`tb- Locate target AD Connector from `"Connectors`" tab."
                $consoleMessage | Write-Host -fore Yellow                

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
                $htmlMessageList.Add($htmlMessage)


                $consoleMessage = "`tc- Locate target AD Connector from `"Connectors`" tab."
                $consoleMessage | Write-Host -fore Yellow

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
                $htmlMessageList.Add($htmlMessage)


                $consoleMessage = "`td- Right click on the AD Connector and go to `"Properties`"."
                $consoleMessage | Write-Host -fore Yellow

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
                $htmlMessageList.Add($htmlMessage)


                $consoleMessage = "`te- Go to `"Select attributes`" option and select `"Show All`" checkbox."
                $consoleMessage | Write-Host -fore Yellow

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
                $htmlMessageList.Add($htmlMessage)


                $consoleMessage = "`tf- Select checkbox for the `"msExchHideFromAddressLists`" attribute."
                $consoleMessage | Write-Host -fore Yellow
                Write-Host "`r`n"

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
                $htmlMessageList.Add($htmlMessage)

                $step++

                #
                # A change is needed to pick the "msExchHideFromAddressLists" attribute in the next import.
                #
                $consoleMessage = "$($step)- Please set `"msExchHideFromAddressLists`" attribute value to at first `"Not set`" and then `"TRUE`" within your on-premises directory. If the value is already `"Not set`", then just change it to `"TRUE`"."
                $consoleMessage | Write-Host -fore White
                Write-Host "`r`n"

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
                $htmlMessageList.Add($htmlMessage)

                $step++    
            }
            elseif ($adObjectAttributeValue -eq $null)
            {
                #
                # Customer needs to set the "msExchHideFromAddressLists" attribute value in the on-premises directory.
                #
                $consoleMessage = "$($step)- Please set `"msExchHideFromAddressLists`" attribute value to `"TRUE`" within your on-premises directory."
                $consoleMessage | Write-Host -fore White
                Write-Host "`r`n"

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
                $htmlMessageList.Add($htmlMessage)

                $step++
            }
            else
            {
                #
                # In spite of the following conditions, "msExchHideFromAddressLists" attribute is NOT populated for the AD connector space object.
                #
                #       + AD Connector schema has "msExchHideFromAddressLists" attribute.
                #
                #       + "msExchHideFromAddressLists" attribute is added to attribute inclusion list of the AD Connector.
                #
                #       + "msExchHideFromAddressLists" attribute of on-premises object is set to some value.
                #
                $consoleMessage = "$($step)- Your settings are correct. Please open a service request through Azure Portal or Office 365 Portal."
                $consoleMessage | Write-Host -fore White
                Write-Host "`r`n"

                $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
                $htmlMessageList.Add($htmlMessage)


                $message = "`tMicrosoft Azure Portal:"
                $consoleMessage = GetConsoleMessageWithLink($message)($global:AzurePortalSupportBladeUrl)
                $consoleMessage | Write-Host -fore Yellow
                Write-Host "`r`n"

                $htmlMessage = WriteHyperlink($global:AzurePortalSupportBladeUrl)($global:AzurePortalSupportBladeText)
                $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
                $htmlMessageList.Add($htmlMessage)


                $message = "`tOffice 365 Portal:"
                $consoleMessage = GetConsoleMessageWithLink($message)($global:OfficePortalUrl)
                $consoleMessage | Write-Host -fore Yellow
                Write-Host "`r`n"

                $htmlMessage = WriteHyperlink($global:OfficePortalUrl)($global:OfficePortalText)
                $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
                $htmlMessageList.Add($htmlMessage)


                $step++
            }
        }
    }

    #
    # If AD connector space object does NOT have "mailNickname" attribute, then it will be out of scope for "In from AD - User Exchange" synchronization rule.
    #
    # As a result, "msExchHideFromAddressLists" attribute will NOT flow into the metaverse.
    #
    if ($mailNickNameValue -eq $null)
    {
        $consoleMessage = "$($step)- Please set the `"mailNickname`" attribute to some value within your on-premises directory."
        $consoleMessage | Write-Host -fore White
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -fontWeight "600" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)
    }

    $consoleMessage = "The account should be hidden from the global address list of the Exchange Online by the end of next sync cycle."
    $consoleMessage | Write-Host -fore Yellow

    $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 1
    $htmlMessageList.Add($htmlMessage)

    
    $hideFromAddressListHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)($global:HideFromGlobalAddressList)

    $SynchronizationIssueHtmlItems.Add($hideFromAddressListHtmlItem)
}

#
# Get if AAD Tenant DirSync Feature "SynchronizeUpnForManagedUsers" enabled or not.
#
# If this feature is DISABLED, then on-prem upn suffix updates will NOT be synchronized to cloud upn for the users
# that are MANAGED and HAS A LICENSE assigned to them.
# 
# Updates to on-prem upn suffix are NOT synchronized to cloud upn for FEDERATED users at all.
#
Function IsSynchronizeUpnForManagedUsersFeatureEnabled
{
    $isSynchronizeUpnForManagedUsersFeatureEnabled = $(Get-MsolDirSyncFeatures -Feature "SynchronizeUpnForManagedUsers").Enabled

    Write-Output $isSynchronizeUpnForManagedUsersFeatureEnabled
}

#
# Get default domain name "<Initial-default-domain-name>.onmicrosoft.com" of the AAD Tenant.
#
Function GetAADTenantDefaultDomainName
{
    param
    (
        [Microsoft.Online.Administration.Domain[]]
        [AllowNull()]
        [parameter(mandatory=$true)]
        $AADTenantDomains
    )

    if ($isNonInteractiveMode)
    {
        return $null
    }

    $defaultDomainName = $null	
    foreach ($domain in $AADTenantDomains)
    {
        if ($domain.Name.Contains(".onmicrosoft.com") -and !$domain.Name.Contains(".mail."))
        {
            $defaultDomainName = $domain.Name
            break
        }
    }

    return $defaultDomainName
}

Function IsUPNSuffixVerifiedInAADTenant
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $upnSuffix,

        [Microsoft.Online.Administration.Domain[]]
        [parameter(mandatory=$true)]
        $AADTenantDomains
    )

    $isUpnSuffixVerified = $false

    foreach ($domain in $AADTenantDomains)
    {
        if ($domain.Name -eq $upnSuffix)
        {
            if ($domain.Status -eq "Verified")
            {
                $isUpnSuffixVerified = $true
            }

            break
        }
    }

    Write-Output $isUpnSuffixVerified
}

#
# Get AAD Tenant user object that corresponds to given metaverse object.
#
Function GetAADTenantUser
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        $MvObject
    )

    if ($isNonInteractiveMode)
    {
        return $null
    }

    $aadTenantUser = $null

    #
    # Try to get AAD Tenant user object by Object Id.
    #
    # AAD Tenant Object Id is a GUID value that is embedded into "cloudAnchor" attribute of the metaverse object.
    #
    $cloudAnchorAttribute = GetMVObjectAttribute($MvObject)("cloudAnchor")

    if ($cloudAnchorAttribute -ne $null)
    {
        $cloudAnchorValue = $cloudAnchorAttribute.Values[0]

        $aadTenantObjectId = $cloudAnchorValue.Split('_')[1]

        $aadTenantUser = Get-MsolUser -ObjectId $aadTenantObjectId

        Write-Output $aadTenantUser

        return
    }

    #
    # In case metaverse object does NOT have "cloudAnchor" attribute, then 
    #
    # AAD Tenant user object will be looked up by "displayName" and "sourceAnchor" attributes of the metaverse object.
    #

    #
    # Get "displayName" attribute value of the metaverse object
    #
    $displayNameAttribute = GetMVObjectAttribute($MvObject)("displayName")
    $displayNameValue = $null

    if ($displayNameAttribute -ne $null)
    {
        $displayNameValue = $displayNameAttribute.Values[0]
    }
    else
    {
        Write-Output $null

        return
    }

    #
    # Get "sourceAnchor" attribute value of the metaverse object
    #
    $sourceAnchorAttribute = GetMVObjectAttribute($MvObject)("sourceAnchor")
    $sourceAnchorValue = $null

    if ($sourceAnchorAttribute -ne $null)
    {
        $sourceAnchorValue = $sourceAnchorAttribute.Values[0]
    }
    else
    {
        Write-Output $null

        return
    }

    #
    # Search user objects in AAD Tenant by "displayName" attribute value of the metaverse object
    #
    $aadTenantUserList = Get-MsolUser -SearchString $displayNameValue

    if ($aadTenantUserList -eq $null)
    {
        Write-Output $null

        return
    }

    #
    # Try to get matching AAD Tenant user object by "sourceAnchor" attribute value of the metaverse object
    #
    $aadTenantUser = $aadTenantUserList | Where-Object{$_.ImmutableId -eq $sourceAnchorValue}

    Write-Output $aadTenantUser
}

#
# Get AAD Tenant group object that corresponds to given metaverse object.
#
Function GetAADTenantGroup
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        $MvObject
    )

    if ($isNonInteractiveMode)
    {
        return $null
    }

    $aadTenantGroup = $null

    #
    # Try to get AAD Tenant group object by Object Id.
    #
    # AAD Tenant Object Id is a GUID value that is embedded into "cloudAnchor" attribute of the metaverse object.
    #
    $cloudAnchorAttribute = GetMVObjectAttribute($MvObject)("cloudAnchor")

    if ($cloudAnchorAttribute -ne $null)
    {
        $cloudAnchorValue = $cloudAnchorAttribute.Values[0]

        $aadTenantObjectId = $cloudAnchorValue.Split('_')[1]

        $aadTenantGroup = Get-MsolGroup -ObjectId $aadTenantObjectId

        Write-Output $aadTenantGroup

        return
    }

    #
    # In case metaverse object does NOT have "cloudAnchor" attribute, then 
    #
    # AAD Tenant group object will be looked up by "displayName" and "sourceAnchor" attributes of the metaverse object.
    #

    #
    # Get "displayName" attribute value of the metaverse object
    #
    $displayNameAttribute = GetMVObjectAttribute($MvObject)("displayName")
    $displayNameValue = $null

    if ($displayNameAttribute -ne $null)
    {
        $displayNameValue = $displayNameAttribute.Values[0]
    }
    else
    {
        Write-Output $null

        return
    }

    #
    # Get "sourceAnchor" attribute value of the metaverse object
    #
    $sourceAnchorAttribute = GetMVObjectAttribute($MvObject)("sourceAnchor")
    $sourceAnchorValue = $null

    if ($sourceAnchorAttribute -ne $null)
    {
        $sourceAnchorValue = $sourceAnchorAttribute.Values[0]
    }
    else
    {
        Write-Output $null

        return
    }

    #
    # Search group objects in AAD Tenant by "displayName" attribute value of the metaverse object
    #
    $aadTenantGroupList = Get-MsolGroup -SearchString $displayNameValue

    if ($aadTenantGroupList -eq $null)
    {
        Write-Output $null

        return
    }

    #
    # Try to get matching AAD Tenant user object by "sourceAnchor" attribute value of the metaverse object
    #
    $aadTenantGroup = $aadTenantGroupList | Where-Object{$_.ImmutableId -eq $sourceAnchorValue}

    Write-Output $aadTenantGroup
}

#
# Get AAD Tenant contact object that corresponds to given metaverse object.
#
Function GetAADTenantContact
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        $MvObject
    )

    if ($isNonInteractiveMode)
    {
        return $null
    }

    $aadTenantContact = $null

    #
    # Try to get AAD Tenant contact object by Object Id.
    #
    # AAD Tenant Object Id is a GUID value that is embedded into "cloudAnchor" attribute of the metaverse object.
    #
    $cloudAnchorAttribute = GetMVObjectAttribute($MvObject)("cloudAnchor")

    if ($cloudAnchorAttribute -ne $null)
    {
        $cloudAnchorValue = $cloudAnchorAttribute.Values[0]

        $aadTenantObjectId = $cloudAnchorValue.Split('_')[1]

        $aadTenantContact = Get-MsolContact -ObjectId $aadTenantObjectId

        Write-Output $aadTenantContact
        return
    }

    #
    # In case metaverse object does NOT have "cloudAnchor" attribute, then 
    #
    # AAD Tenant contact object will be looked up by "displayName" and "sourceAnchor" attributes of the metaverse object.
    #

    #
    # Get "displayName" attribute value of the metaverse object
    #
    $displayNameAttribute = GetMVObjectAttribute($MvObject)("displayName")
    $displayNameValue = $null

    if ($displayNameAttribute -ne $null)
    {
        $displayNameValue = $displayNameAttribute.Values[0]
    }
    else
    {
        Write-Output $null
        return
    }

    #
    # Get "sourceAnchor" attribute value of the metaverse object
    #
    $sourceAnchorAttribute = GetMVObjectAttribute($MvObject)("sourceAnchor")
    $sourceAnchorValue = $null

    if ($sourceAnchorAttribute -ne $null)
    {
        $sourceAnchorValue = $sourceAnchorAttribute.Values[0]
    }
    else
    {
        Write-Output $null
        return
    }

    #
    # Search contact objects in AAD Tenant by "displayName" attribute value of the metaverse object
    #
    $aadTenantContactList = Get-MsolContact -SearchString $displayNameValue

    if ($aadTenantContactList -eq $null)
    {
        Write-Output $null
        return
    }

    #
    # Try to get matching AAD Tenant object by "sourceAnchor" attribute value of the metaverse object
    #
    $aadTenantContact = $aadTenantContactList | Where-Object{$_.ImmutableId -eq $sourceAnchorValue}
    Write-Output $aadTenantContact
}

Function GetCSObjectAttribute
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        $CsObject,

        [string]
        [parameter(mandatory=$true)]
        $AttributeName
    )

    if ($CsObject.Attributes[$AttributeName] -eq $null)
    {
        WriteEventLog($EventIdCsObjectAttributeNotFound)($EventMsgCsObjectAttributeNotFound -f $AttributeName)

        Write-Output $null
    }
    else
    {
        Write-Output $CsObject.Attributes[$AttributeName]
    }
}

Function GetMVObjectAttribute
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        $MvObject,

        [string]
        [parameter(mandatory=$true)]
        $AttributeName
    )

    if ($MvObject.Attributes[$AttributeName] -eq $null)
    {
        WriteEventLog($EventIdMVObjectAttributeNotFound)($EventMsgMVObjectAttributeNotFound -f $AttributeName)

        Write-Output $null
    }
    else
    {
        Write-Output $MvObject.Attributes[$AttributeName];
    }
}

# Set/Create OutputDirectory global variable
Function Set-OutputDirectory
{
    if(-not (Test-Path $ObjectDiagnosticsReportOutputDirectory))  
    {
        $folder = New-Item -Path $ObjectDiagnosticsReportOutputDirectory -ItemType directory
    }
}

Function Export-ObjectDiagnosticsHtmlReport  
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $Title,

        [string]
        [parameter(mandatory=$true)]
        $ReportDate,

        [string]
        [parameter(mandatory=$true)]
        $HtmlDoc
    )

    if ($isNonInteractiveMode)
    {
        return
    }
    
    $filename = $Title.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'

    $filename = "$ObjectDiagnosticsReportOutputDirectory\$ReportDate-$filename.htm"
    
    try 
    {
        $HtmlDoc | Out-File -FilePath $filename

        "HTML REPORT:" | Write-Host -fore Cyan
        "------------" | Write-Host -fore Cyan

        "Detailed HTML Report has been exported to file $filename." | Write-Host -fore Green

        Write-Host "`r`n"
    }
    catch
    {
        Write-Error "An error occurred while exporting HTML report to $filename : $($_.Exception.Message)"

        return
    }

    "Opening the html report in Internet Explorer..." | Write-Host -fore White
    
    try
    {
        #
        # Add 'Microsoft.VisualBasic' namespace into PowerShell session.
        #
        Add-Type -AssemblyName "Microsoft.VisualBasic"
        
        $internetExplorer = New-Object -com internetexplorer.application

        $internetExplorer.navigate2($filename)
        $internetExplorer.visible = $true

        if ($internetExplorer.Busy)
        {
            Sleep -Seconds 15
        }

        $ieProcess = Get-Process | ? { $_.MainWindowHandle -eq $internetExplorer.HWND }

        #
        # Set focus to Internet Explorer so that it will appear on top of other windows.
        #
        [Microsoft.VisualBasic.Interaction]::AppActivate($ieProcess.Id)
    }
    catch
    {
        Write-Error "Unable to open Internet Explorer : $($_.Exception.Message)"
    }
}

Function CheckAttributeBasedFiltering
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $ADConnectorName,

        [string]
        [parameter(mandatory=$true)]
        $ObjectDN,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$false)]
        [AllowNull()]
        $AdCsObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$false)]
        [AllowNull()]
        $MvObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$false)]
        [AllowNull()]
        $AadCsObject
    )

    "Checking Attribute Filtering configuration..." | Write-Host -fore White
    
    if ($AdCsObject -ne $null -and ($MvObject -eq $null -or $AadCsObject -eq $null))
    {
        $previewResult = Sync-ADSyncCsObject -ConnectorName $ADConnectorName -DistinguishedName $ObjectDN
        if ($previewResult -ne $null)
        {
            $previewDiagnosticsData = $previewResult.PreviewDiagnosticsData
            if ($previewDiagnosticsData -ne $null)
            {
                foreach ($entryDiagnoticData in $previewDiagnosticsData.EntryModificationDiagnosticsDataList)
                {
                    $scopeModuleDiagnosticsData = $entryDiagnoticData.ScopeModuleDiagnosticsData
                    foreach ($outOfScopeSyncRule in $scopeModuleDiagnosticsData.OutOfScopeSyncRules)
                    {
                        $syncRuleName = $outOfScopeSyncRule.SyncRuleName
                        if ($outOfScopeSyncRule.SourceObjectMarkedForDeletion)
                        {
                            $consoleMessage = "Object `"$ObjectDN`" is marked for deletion. ALL sync rules are out of scope for this object."
                            $consoleMessage | ReportError
                            Write-Host "`r`n"
                        }
                        elseif ($outOfScopeSyncRule.Disabled)
                        {
                            $consoleMessage = "Sync Rule `"$syncRuleName`" is out of scope because it is disabled."
                            $consoleMessage | ReportWarning
                            Write-Host "`r`n"
                        }
                        else
                        {
                            "Sync Rule `"$syncRuleName`" is out of scope because the following scoping conditions were not satisfied:" | ReportWarning
                            
                            $scopeConditionGroupIndex = 1
                            foreach ($scopeConditionGroup in $outOfScopeSyncRule.ScopeConditionGroups)
                            {
                                $attribute = $scopeConditionGroup.Attribute
                                $operator = $scopeConditionGroup.ComparisonOperator
                                $value = $scopeConditionGroup.ComparisonValue
                                "Scoping Group `"$scopeConditionGroupIndex`"" | ReportWarning								
                                "[`"$attribute $operator $value`"]" | ReportWarning
                                $scopeConditionGroupIndex++				
                            }
                            
                            Write-Host "`r`n"
                        }
                    }
                }
            }
        }
    }
    else
    {
        "OK" | Write-Host -fore Green
        Write-Host "`r`n"
    }
}

Function CheckOUBasedFiltering
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.ConnectorPartition]
        [parameter(mandatory=$false)]
        $Partition,

        [string]
        [parameter(mandatory=$true)]
        $ObjectDN
    )

    if ($Partition -eq $null)
    {
        return
    }

    $htmlMessageList = New-Object System.Collections.Generic.List[string]

    "Checking OU Filtering configuration..." | Write-Host
    $containerInclusionList = $Partition.ConnectorPartitionScope.ContainerInclusionList
    $containerExclusionList = $Partition.ConnectorPartitionScope.ContainerExclusionList

    $mostSpecificOUInInclusionList = Get-MostSpecificContainerFromContainerList $containerInclusionList $ObjectDN
    $mostSpecificOUInExclusionList = Get-MostSpecificContainerFromContainerList $containerExclusionList $ObjectDN

    $objectOutOfSyncScope = $false
    if ($mostSpecificOUInInclusionList -ne $null -and $mostSpecificOUInExclusionList -ne $null)
    {
        if ($mostSpecificOUInInclusionList.Length -lt $mostSpecificOUInExclusionList.Length)
        {
            $objectOutOfSyncScope = $true
        }
    }
    elseif ($mostSpecificOUInExclusionList -ne $null)
    {
        $objectOutOfSyncScope = $true
    }

    if ($objectOutOfSyncScope)
    {
        WriteEventLog($EventIdOUFiltered)($EventMsgOUFiltered -f ($ObjectDN, $mostSpecificOUInExclusionList))

        $SynchronizationIssueList.Add($global:OuFilteringIssue)

        Write-Host "`r`n"
        
        "OU FILTERING - ANALYSIS:" | Write-Host -fore Cyan
        "------------------------" | Write-Host -fore Cyan
        
        $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)


        $consoleMessage = "Object `"$ObjectDN`" is not present in sync scope. It belongs to a container $mostSpecificOUInExclusionList that is excluded from syncing."
        $consoleMessage | ReportError
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        "OU FILTERING - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
        "-----------------------------------" | Write-Host -fore Cyan
        
        $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)
        

        $message = "Include the container $mostSpecificOUInExclusionList in the list of organizational units that should be synced. To read more on how to do this, please see:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:OuBasedFilteringUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:OuBasedFilteringUrl)($global:OuBasedFilteringText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 0
        $htmlMessageList.Add($htmlMessage)


        $ouFilteringHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("OU Filtering")

        $SynchronizationIssueHtmlItems.Add($ouFilteringHtmlItem)
    }
    else
    {
        "There is no OU filtering configuration that prevents this object from being imported in the AD connector space." | Write-Host
        Write-Host "`r`n"
    }

    ReportOutput -PropertyName "OU Filtered" -PropertyValue $objectOutOfSyncScope
}

# This function will return the nearest ancestor to the object
# which is present in the specified container list.
# Example: if container list contains "DC=msft,DC=com" and "OU=Sales,DC=msft,DC=com" and objectDN = "CN=Aditis,OU=Sales,DC=msft,DC=com"
# then this function will return "OU=Sales,DC=msft,DC=com"
Function Get-MostSpecificContainerFromContainerList
{
    param
    (
        [string[]]
        [parameter(mandatory=$false)]
        $ContainerList,

        [string]
        [parameter(mandatory=$true)]
        $ObjectDN
    )

    if ($ContainerList -eq $null)
    {
        return
    }

    $length = 0
    $mostSpecificOU = $null
    foreach ($container in $ContainerList)
    {
        if ($ObjectDN.Contains($container))
        {
            if ($container.Length -gt $length)
            {
                $mostSpecificOU = $container
                $length = $container.Length
            }
        }
    }

    return $mostSpecificOU
}

#
# Get if AD object type matches the given object type.
#
Function IsObjectTypeMatch
{
    param
    (
        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        $AdObject,

        [string]
        [parameter(mandatory=$true)]
        $ObjectType
    )

    #
    # AD object attribute "objectclass" is a multi-valued attribute.
    #
    $objectClasses = [System.Collections.ArrayList] $AdObject["objectclass"]

    foreach ($objectClass in $objectClasses)
    {
        if ($objectClass -eq $ObjectType)
        {
            Write-Output $true

            return
        }
    }

    Write-Output $false
}

#
# Check if object type is included on connector
#
Function CheckObjectTypeInclusion
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector]
        [parameter(mandatory=$true)]
        $AdConnector,

        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        $AdObject
    )

    $htmlMessageList = New-Object System.Collections.Generic.List[string]

    Write-Host "Checking if object type is in the connector's object type inclusion list..."

    $connectorClasses = [System.Collections.ArrayList] $AdConnector.ObjectInclusionList
    $objectClasses = [System.Collections.ArrayList] $AdObject["objectclass"]

    $objectClass = $objectClasses[$objectClasses.Count - 1]

    if ($connectorClasses -contains $objectClass)
    {
        Write-Host "Object type `"$objectClass`" is in the connector's object type inclusion list. This configuration will not prevent this object from being imported into the AD connector space."
        ReportOutput -PropertyName "Object Type Inclusion" -PropertyValue "False"
    }
    else
    {
        WriteEventLog($EventIdObjectTypeInclusion)($EventMsgObjectTypeInclusion -f ([String] $AdObject["distinguishedname"], $objectClass, $AdConnector.Identifier))

        $SynchronizationIssueList.Add($global:ObjectTypeInclusionIssue)

        Write-Host "`r`n"
        "OBJECT TYPE INCLUSION - ANALYSIS:" | Write-Host -fore Cyan
        "---------------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        $consoleMessage = "Object is not present in sync scope as it is of object class `"$objectClass`" which is not part of the connector's object type inclusion list." 
        $consoleMessage | ReportError -PropertyName "Object Type Inclusion" -PropertyValue "True"
        Write-Host "`r`n"

        $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)


        "OBJECT TYPE INCLUSION - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
        "--------------------------------------------" | Write-Host -fore Cyan

        $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
        $htmlMessageList.Add($htmlMessage)

        $message = "Ensure object type `"$objectClass`" is being used in a sync rule for this connector. For more information on sync rules, please see:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:CustomizeSyncRulesUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:CustomizeSyncRulesUrl)($global:CustomizeSyncRulesText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 2
        $htmlMessageList.Add($htmlMessage)

        $message = "If object type `"$objectClass`" is not available in the Synchronization Rules Editor, use the Wizard to refresh the directory schema. For more information on refreshing the schema, please see:"
        $consoleMessage = GetConsoleMessageWithLink($message)($global:RefreshSchemaUrl)
        $consoleMessage | Write-Host -fore Yellow
        Write-Host "`r`n"

        $htmlMessage = GetHtmlMessageWithLink($message)($global:RefreshSchemaUrl)($global:RefreshSchemaText)
        $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 0
        $htmlMessageList.Add($htmlMessage)


        $objectInclusionHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("Object Type Inclusion")

        $SynchronizationIssueHtmlItems.Add($objectInclusionHtmlItem)

        "Object type `"$objectClass`" is not in the connector's object type inclusion list. This will prevent this object from being imported into the AD connector space." | ReportError
    }

    Write-Host "`r`n"
}

#
# Check group filtering
#
Function CheckGroupFiltering
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector]
        [parameter(mandatory=$true)]
        $AdConnector,

        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        $AdObject
    )

    $htmlMessageList = New-Object System.Collections.Generic.List[string]

    Write-Host "Checking group filtering configuration..."

    $globalSettings = Get-ADSyncGlobalSettings
    $groupFilteringDN = $AdConnector.GlobalParameters["Connector.GroupFilteringGroupDn"].Value

    if (($globalSettings.Parameters["Microsoft.OptionalFeature.GroupFiltering"].Value -eq "True") -and $groupFilteringDN)
    {
        Write-Host "Group filtering is enabled for group `"$groupFilteringDN`"."

        $groupADobject = Search-ADSyncDirectoryObjects -AdConnectorId $adConnector.Identifier -LdapFilter "(distinguishedName=$groupFilteringDN)" -SearchScope Subtree -SizeLimit 1

        $adObjectCSobject = GetCSObject($AdConnector.Name)($AdObject["distinguishedname"])
        $groupCSobject = GetCSObject($AdConnector.Name)($groupFilteringDN)

        $isMemberInAD = IsMemberOfGroupInAD($groupADobject[0])([String] $AdObject["distinguishedname"])
        $isMemberInCS = IsMemberOfGroupInCS($groupCSobject)($adObjectCSobject)

        if ($isMemberInAD -and $isMemberInCS)
        {
            Write-Host "No issues detected. The object `"$([String] $AdObject["distinguishedname"])`" is a member of group `"$groupFilteringDN`" that is configured for group filtering."
            ReportOutput -PropertyName "Group Filtered" -PropertyValue "False"
        }
        else
        {
            WriteEventLog($EventIdGroupFiltered)($EventMsgGroupFiltered -f ([String] $AdObject["distinguishedname"], $AdConnector.Identifier, $groupFilteringDN))

            $SynchronizationIssueList.Add($global:GroupFilteringIssue)

            Write-Host "`r`n"
            "GROUP FILTERING - ANALYSIS:" | Write-Host -fore Cyan
            "---------------------------" | Write-Host -fore Cyan

            $htmlMessage = WriteHtmlMessage -message "Analysis:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
            $htmlMessageList.Add($htmlMessage)

            if (!$isMemberInAD)
            {
                $consoleMessage = "Object is not present in sync scope as group filtering is enabled and the object is not a member of filtering group `"$groupFilteringDN`"." 
                $consoleMessage | ReportError -PropertyName "Group Filtered" -PropertyValue "True"
            }
            else
            {
                $consoleMessage = "Object is a member of the filtering group `"$groupFilteringDN`" in the Active Directory but is not a member of the group in the AD Connector Space." 
                $consoleMessage | ReportError -PropertyName "Group Filtered" -PropertyValue "True"
            }
            
            Write-Host "`r`n"

            $htmlMessage = WriteHtmlMessage -message $consoleMessage -color "#252525" -numberOfLineBreaks 2
            $htmlMessageList.Add($htmlMessage)


            "GROUP FILTERING - RECOMMENDED ACTIONS:" | Write-Host -fore Cyan
            "--------------------------------------" | Write-Host -fore Cyan

            $htmlMessage = WriteHtmlMessage -message "Recommended Actions:" -color "#252525" -fontSize "16px" -fontWeight "700" -numberOfLineBreaks 1
            $htmlMessageList.Add($htmlMessage)

            if (!$isMemberInAD)
            {
                $message = "Add the object to group `"$groupFilteringDN`" or change group filtering settings. For more information on group filtering, please see:"
                $consoleMessage = GetConsoleMessageWithLink($message)($global:ConfigureGroupSyncFilteringUrl)
                $consoleMessage | Write-Host -fore Yellow

                $htmlMessage = GetHtmlMessageWithLink($message)($global:ConfigureGroupSyncFilteringUrl)($global:ConfigureGroupSyncFilteringText)
            }
            else
            {
                $message = "Ensure an import operation has been run on the AD Connector since the object was added to group `"$groupFilteringDN`" and that there were no errors during the import. More information on operations can be found here:"
                $consoleMessage = GetConsoleMessageWithLink($message)($global:OperationsTabUrl)
                $consoleMessage | Write-Host -fore Yellow

                $htmlMessage = GetHtmlMessageWithLink($message)($global:OperationsTabUrl)($global:OperationsTabText)
            }

            Write-Host "`r`n"

            $htmlMessage = WriteHtmlMessage -message $htmlMessage -color "#252525" -numberOfLineBreaks 0
            $htmlMessageList.Add($htmlMessage)


            $groupFilteringHtmlItem = WriteHtmlAccordionItemForParagraph($htmlMessageList)("Group Filtering")

            $SynchronizationIssueHtmlItems.Add($groupFilteringHtmlItem)
        }
    }
    else
    {
        Write-Host "No issues detected - group filtering is not in use."
        ReportOutput -PropertyName "Group Filtered" -PropertyValue "False"
    }

    Write-Host "`r`n"
}

Function IsMemberOfGroupInAD
{
    param
    (
        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        $AdObject,

        [string]
        [parameter(mandatory=$true)]
        $MemberDN
    )

    #
    # AD object attribute "member" is a multi-valued attribute.
    #
    $members = [System.Collections.ArrayList] $AdObject["member"]

    foreach ($member in $members)
    {
        if ($member -eq $MemberDN)
        {
            return $true
        }
    }

    return $false
}

Function IsMemberOfGroupInCS
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdCsGroupObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdCsMemberObject
    )

    if ($AdCsGroupObject -ne $null -and $AdCsMemberObject -ne $null)
    {
        $members = [System.Collections.Generic.List[String]]$AdCsGroupObject.Attributes["member"].Values
        foreach ($member in $members)
        {
            if ($member -eq $AdCsMemberObject.DistinguishedName)
            {
                return $true
            }
        }
    }

    return $false
}

Function IsMemberOfGroupInMV
{
    param
    (
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $MvGroupObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $MvMemberObject
    )

    if ($MvGroupObject -ne $null -and $MvMemberObject -ne $null)
    {
        $memberToCheck = [System.Guid]::Parse($MvMemberObject.ObjectId)
        $members = [System.Collections.Generic.List[String]]$MvGroupObject.Attributes["member"].Values
        foreach ($member in $members)
        {
            $actualMember = [System.Guid]::Parse($member)
            if ($actualMember.Equals($memberToCheck))
            {
                return $true
            }
        }
    }

    return $false
}

Function GetObjectAllADAttributeHtmlGroup
{
    param
    (
        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        $AdObjectFromConnectorAccount,

        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        $AdObjectFromProvidedAccount,

        [String]
        [parameter(mandatory=$true)]
        $ConnectorAccountName,

        [String]
        [parameter(mandatory=$true)]
        $ProvidedAccountName
    )

    $htmlItems = @()

    $tableHeaders = ("Attribute Name", "Attribute Value retrieved by $ProvidedAccountName", "Attribute Value retrieved by $ConnectorAccountName")

    #
    # Get all Attributes from objects retrieved by both the connector and provided account
    #
    $ProvidedAccountAttributesHashTable = ConvertADObjectToHashTable($AdObjectFromProvidedAccount)($true)
    $ConnectorAccountAttributesHashTable = ConvertADObjectToHashTable($AdObjectFromConnectorAccount)($true)

    $attributesCompareHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlADAttributesComparisonTitle)("ADObjectAttributeComparisonTable")($tableHeaders)($ProvidedAccountAttributesHashTable)($global:HtmlADObjectType)($ConnectorAccountAttributesHashTable)

    $htmlItems += $attributesCompareHtmlItem

    if ($htmlItems.length -gt 0)
    {
        $htmlGroup = WriteHtmlAccordionGroup($htmlItems)($global:HtmlAttributeDetailsSectionTitle)
    }

    Write-Output $htmlGroup
}

Function GetUserObjectHtmlGroup
{
    param
    (
        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdCsObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $MvObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadCsObject,

        [Microsoft.Online.Administration.User]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadUserObject
    )

    $htmlItems = @()

    $tableHeaders = ("Attribute Name", "Attribute Value")

    #
    # On-Premises AD Object
    #
    if ($AdObject)
    {
        $AdObjectHashTable = ConvertADObjectToHashTable($AdObject)
    
        $AdObjectHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlADObjectTitle)("ADObjectTable")($tableHeaders)($AdObjectHashTable)($global:HtmlADObjectType)

        $htmlItems += $AdObjectHtmlItem
    }

    #
    # Metaverse Object
    #
    if ($MvObject)
    {
        $MvObjectHashTable = ConvertMVObjectToHashTable($MvObject)

        $MvObjectHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlAADConnectObjectTitle)("AADConnectObjectTable")($tableHeaders)($MvObjectHashTable)($global:HtmlAADConnectObjectType)

        $htmlItems += $MvObjectHtmlItem
    }

    #
    # Azure AD Object
    #
    if ($AadUserObject)
    {
        $AadUserObjectHashTable = ConvertAADUserObjectToHashTable($AadUserObject)
    
        $AadUserObjectHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlAzureADObjectTitle)("AADObjectTable")($tableHeaders)($AadUserObjectHashTable)($global:HtmlAzureADObjectType)
    
        $htmlItems += $AadUserObjectHtmlItem
    }
    
    $htmlGroup = $null

    if ($htmlItems.length -gt 0)
    {
        $htmlGroup = WriteHtmlAccordionGroup($htmlItems)($global:HtmlObjectDetailsSectionTitle)
    }

    Write-Output $htmlGroup
}

Function GetGroupObjectHtmlGroup
{
    param
    (
        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdCsObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $MvObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadCsObject,

        [Microsoft.Online.Administration.Group]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadGroupObject
    )

    $htmlItems = @()

    $tableHeaders = ("Attribute Name", "Attribute Value")

    #
    # On-Premises AD Object
    #
    if ($AdObject)
    {
        $AdObjectHashTable = ConvertADObjectToHashTable($AdObject)
    
        $AdObjectHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlADObjectTitle)("ADObjectTable")($tableHeaders)($AdObjectHashTable)($global:HtmlADObjectType)

        $htmlItems += $AdObjectHtmlItem
    }

    #
    # Metaverse Object
    #
    if ($MvObject)
    {
        $MvObjectHashTable = ConvertMVObjectToHashTable($MvObject)

        $MvObjectHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlAADConnectObjectTitle)("AADConnectObjectTable")($tableHeaders)($MvObjectHashTable)($global:HtmlAADConnectObjectType)

        $htmlItems += $MvObjectHtmlItem
    }

    #
    # Azure AD Object
    #
    if ($AadGroupObject)
    {
        $AadGroupObjectHashTable = ConvertAADGroupObjectToHashTable($AadGroupObject)
    
        $AadGroupObjectHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlAzureADObjectTitle)("AADObjectTable")($tableHeaders)($AadGroupObjectHashTable)($global:HtmlAzureADObjectType)
    
        $htmlItems += $AadGroupObjectHtmlItem
    }
    
    $htmlGroup = $null

    if ($htmlItems.length -gt 0)
    {
        $htmlGroup = WriteHtmlAccordionGroup($htmlItems)($global:HtmlObjectDetailsSectionTitle)
    }

    Write-Output $htmlGroup
}

Function GetContactObjectHtmlGroup
{
    param
    (
        [System.Collections.Generic.Dictionary[[String], [Object]]]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AdCsObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.MvObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $MvObject,

        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadCsObject,

        [Microsoft.Online.Administration.Contact]
        [parameter(mandatory=$true)]
        [AllowNull()]
        $AadContactObject
    )

    $htmlItems = @()

    $tableHeaders = ("Attribute Name", "Attribute Value")

    #
    # On-Premises AD Object
    #
    if ($AdObject)
    {
        $AdObjectHashTable = ConvertADObjectToHashTable($AdObject)
    
        $AdObjectHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlADObjectTitle)("ADObjectTable")($tableHeaders)($AdObjectHashTable)($global:HtmlADObjectType)

        $htmlItems += $AdObjectHtmlItem
    }

    #
    # Metaverse Object
    #
    if ($MvObject)
    {
        $MvObjectHashTable = ConvertMVObjectToHashTable($MvObject)

        $MvObjectHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlAADConnectObjectTitle)("AADConnectObjectTable")($tableHeaders)($MvObjectHashTable)($global:HtmlAADConnectObjectType)

        $htmlItems += $MvObjectHtmlItem
    }

    #
    # Azure AD Object
    #
    if ($AadContactObject)
    {
        $AadContactObjectHashTable = ConvertAADContactObjectToHashTable($AadContactObject)
    
        $AadContactObjectHtmlItem = WriteHtmlAccordionItemForTable($global:HtmlAzureADObjectTitle)("AADObjectTable")($tableHeaders)($AadContactObjectHashTable)($global:HtmlAzureADObjectType)
    
        $htmlItems += $AadContactObjectHtmlItem
    }
    
    $htmlGroup = $null

    if ($htmlItems.length -gt 0)
    {
        $htmlGroup = WriteHtmlAccordionGroup($htmlItems)($global:HtmlObjectDetailsSectionTitle)
    }

    Write-Output $htmlGroup
}

Function GetConsoleMessageWithLink
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $message,

        [string]
        [parameter(mandatory=$true)]
        $url
    )

    $consoleMessage = $message
    $consoleMessage += " "
    $consoleMessage += $url

    Write-Output $consoleMessage
}

Function GetHtmlMessageWithLink
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $message,

        [string]
        [parameter(mandatory=$true)]
        $url,

        [string]
        [parameter(mandatory=$true)]
        $text
    )

    $hyperlink = WriteHyperlink($url)($text)

    $htmlMessage = $message
    $htmlMessage += " "
    $htmlMessage += $hyperlink

    Write-Output $htmlMessage
}

Function AskIfToolHelpful
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $objectDN
    )

    if ($isNonInteractiveMode)
    {
        return
    }

    foreach ($synchronizationIssue in $SynchronizationIssueList)
    {
        do
        {
            $answer = Read-Host "Did you find this tool helpful about the `"$($synchronizationIssue)`" issue? [y/n]"
        } 
        while(($answer -ne 'y') -and ($answer -ne 'Y') -and ($answer -ne 'n') -and ($answer -ne 'N'))

        $eventMessage = "Object: "
        $eventMessage += $objectDN
        $eventMessage += "`n"

        $eventMessage += "Synchronization Issue: "
        $eventMessage += $synchronizationIssue
        $eventMessage += "`n"
        
        $eventMessage += "Is Tool Helpful: "
        $eventMessage += $answer

        WriteEventLog($EventIdIsToolHelpful)($eventMessage)

        Write-Host "`r`n"
    }
}

Function WriteTitle
{
    param
    (
        [string]
        [parameter(mandatory=$true)]
        $title
    )

    $lineSizeWithoutText = 50

    $lineSize = $title.Length + $lineSizeWithoutText

    #
    # Border line - top and bottom of the title
    #
    for ($i = 0; $i -lt $lineSize; $i++)
    {
        $borderLine += '='
    }

    #
    # Middle line without text
    #
    $midLine = '='

    for ($i = 0; $i -lt $lineSize-2; $i++)
    {
        $midLine += ' '
    }
    
    $midLine += '='

    #
    # Title line
    #
    $titleLine = '='

    for ($i = 0; $i -lt ($lineSizeWithoutText-2)/2; $i++)
    {
        $titleLine += ' '
    }

    $titleLine += $title

    for ($i = 0; $i -lt ($lineSizeWithoutText-2)/2; $i++)
    {
        $titleLine += ' '
    }

    $titleLine += '='

    #
    # Resulting Title
    #
    Write-Host $borderLine
    Write-Host $midLine
    Write-Host $titleLine
    Write-Host $midLine
    Write-Host $borderLine
}
# SIG # Begin signature block
# MIIoNgYJKoZIhvcNAQcCoIIoJzCCKCMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC0H9lnZ1hwvjvm
# sdKdzB3OUcmzXO/jyfBBpnMZZgRKvqCCDYIwggYAMIID6KADAgECAhMzAAADXJXz
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJwuY4Do
# oqbDuVQhI1poQ+6ALeONOzCIN7+typiYrhFzMEIGCisGAQQBgjcCAQwxNDAyoBSA
# EgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20w
# DQYJKoZIhvcNAQEBBQAEggEAomik0BRSUlsWpp7ZS3TSMGUEetp/HYxSCj0Q8M9B
# hUOa1TY09z58VUDCBImCpyFl3LJda2PdHk9ERX0BFC0A+GumGXXRrRtfpV+ebAzB
# SL96Ocd3UnlkBgwfPYDTyHnQfjv9TGn2xe6TM1tOVKQhab5GtM3bttJzgtGMDvD3
# 6K+r1OnAsGKPMGf+6tebIpb6+VCzfMCHu+0EdLQ9RvtLswUdcFWBYqA95/04NDdJ
# 1JvoIS4pymfO7NNNJSZlO7aPsXySqctlgSEmWl1U9clpaM7aTqKeTVbUC1jmoTwZ
# ebN6rzWz9PLOKUrLHSDmEyo7fwebRKeQWZVfpl4jiZWo3qGCF5QwgheQBgorBgEE
# AYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUD
# BAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoD
# ATAxMA0GCWCGSAFlAwQCAQUABCDA7BLvqJ7GfO4al0vyBFzYCCcaVh6jIDAD/biq
# etbsrgIGZQQ0jjx0GBMyMDIzMTAwNDE5MjgyMC41NjVaMASAAgH0oIHRpIHOMIHL
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
# SIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCAb2sqF64tfVHwL0KI8jyicYVMc5JDW
# 1swwfelnWpBEFDCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIDJsz0F6L0XD
# Um53JRBfNMZszKsllLDMiaFZ3PL/LqxnMIGYMIGApH4wfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAHODxj3RZfnxv8AAQAAAc4wIgQg1+caYgNLKhNy
# gXcmWt0L0jo/QJsgWkeqnT9LPZdM47kwDQYJKoZIhvcNAQELBQAEggIAr+1tc7BQ
# fsVmbVsomvxysGFzHhzs/HUHwixUo1hbl9rrxWjYuUn2/yFyVsLU0VlZNx/ISCuZ
# 1/H0zotzBBQKvAQkHqbRdVnfiSQ/4iQ8fXPBtEAvnqMsiCa4uTZQp+9nXKzDLH7t
# 0K4kncB4FbOssUS2dWO713SZiONhsaifxYqG4+h+R7zvwvxkB79cekJbVUeFDu6Q
# RuIYO/VrX6w1L3au2+DmBCOLdak6DQCt6S0QY3+GElR7dAnE+4I4usieUJSzZD7d
# 1Wu435zAxL2eLZY4iM2HvXfb60Ef+aHc4KPWgHOZaDQ7HDu7iD0vGB5wk6aRf5LQ
# fb4YGWFWar+tuiA4YvHVKBhJos5Nc1PIzFiVrX9uLaFJgRZjsu+nWFYou6bq4+yy
# WjyganTL6dG10HKV4jeiPX8DTs93jN7aU+0a8hSqy/NbAGOD9be4x2UPx03xj/lf
# vZZ7fQDrcqYkvu8/6Gp0XYJF9uDO+eE5tKUjw4HnMxRO8znnyhXyBsy00nUSKCPl
# qYH8lLL9eM6C8ZEiPJuNrLAnrSdbrsmY3uFn0+f8+JuW1TzzPvvXUUkKk5KvhzSo
# YRdeKtegAs5Pkf93rjojzO71BuUtJDksC2mMhEjh1WCXVev6GWVCPF78w9ghEul9
# kJxOFvhK2rlhKkx6cfeN4ZoVQcuwJbi1IQE=
# SIG # End signature block
