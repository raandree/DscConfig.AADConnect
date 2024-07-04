#-------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation.  All rights reserved.
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
#region script variables
#-------------------------------------------------------------------------

$Script:ConfigureAdDsConnectorAccountLink = "https://aka.ms/aadc-configure-ad-ds-connector-account"

$Script:CustomSyncRuleLink = "https://aka.ms/aadc-custom-sync-rule"

$Script:RefreshDirectorySchemaLink = "https://aka.ms/aadc-refresh-directory-schema"

$Script:SyncServiceManagerUiOperationsLink = "https://aka.ms/aadc-sync-service-manager-ui-operations"

$Script:ExportErrorAttributeValueMustBeUniqueLink = "https://aka.ms/aadc-exporterror-attributevaluemustbeunique"

$Script:ExportErrorDataValidationFailedLink = "https://aka.ms/aadc-exporterror-datavalidationfailed"

$Script:ExportErrorFederatedDomainChangeErrorLink = "https://aka.ms/aadc-exporterror-federateddomainchangeerror"

$Script:ExportErrorInvalidSoftMatchLink = "https://aka.ms/aadc-exporterror-invalidsoftmatch"

$Script:ExportErrorLargeObjectLink = "https://aka.ms/aadc-exporterror-largeobject"

$Script:ExportErrorObjectTypeMismatchLink = "https://aka.ms/aadc-exporterror-objecttypemismatch"

$Script:FilteringDomainLink = "https://aka.ms/aadc-filtering-domain"

$Script:FilteringOrganizationalUnitLink = "https://aka.ms/aadc-filtering-organizational-unit"

$Script:FilteringGroupLink = "https://aka.ms/aadc-filtering-group"

$Script:FilteringAttributeLink = "https://aka.ms/aadc-filtering-attribute"

$Script:SchedulerDisableLink = "https://aka.ms/aadc-scheduler-disable"

$Script:SchedulerStopLink = "https://aka.ms/aadc-scheduler-stop"

$Script:ErrorsLink = "https://aka.ms/aadc-errors"

$Script:GenericRecommendedAction = "Please ensure Azure AD Connect sync engine is not running any operation and disable the scheduler temporarily while using `"Invoke-ADSyncSingleObjectSync`". To learn more on how to disable the scheduler, please see: $Script:SchedulerDisableLink . To learn more on how to stop a synchronization cycle, please see: $Script:SchedulerStopLink . To learn more about error codes, please see: $Script:ErrorsLink . Retry running the `"Invoke-ADSyncSingleObjectSync`" cmdlet with the scheduler disabled and when the sync engine is not running any operation. If the issue persists, open a support case for this issue through the Azure Portal - this will allow a Support Engineer to work with you directly to gather more context and investigate the problem."

#-------------------------------------------------------------------------
#endregion script variables
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
#region class definitions
#-------------------------------------------------------------------------

class AttributeInfo
{
    [string] $Name

    [bool] $IsMultiValued

    [string] $Type

    [string] $Value

    [string] $Add

    [string] $Delete

    [string] $Operation

    [string] $SyncRule

    [string] $MappingType

    [string] $DataSource
}

class ProvisioningSystem
{
    [ValidateSet("Active Directory", "Azure Active Directory")]
    [string] $Name
}

class ProvisioningIdentity
{
    [string] $Id = [string]::Empty

    [string] $Type = [string]::Empty

    [string] $Name = [string]::Empty
}

class StatusInfo
{
    [ValidateSet("Success", "Failure", "Skipped")]
    [string] $Status

    [string] $ErrorCode = [string]::Empty

    [string] $Reason = [string]::Empty

    [string] $AdditionalDetails = [string]::Empty

    [string] $ErrorCategory = [string]::Empty

    [string] $RecommendedAction = [string]::Empty
}

class ProvisioningProperty
{
    [string] $Name

    [string] $OldValue

    [string] $NewValue
}

class ProvisioningStep
{
    [ValidateSet("Success", "Failure", "Skipped")]
    [string] $Status

    [ValidateSet("Scoping", "Import", "Sync", "Export")]
    [string] $Type

    [string] $Name = [string]::Empty

    [string] $Description = [string]::Empty

    [string] $Timestamp = [DateTime]::UtcNow.ToUniversalTime()

    [System.Collections.Generic.Dictionary[[string], [string]]] $Details = [System.Collections.Generic.Dictionary[[string], [string]]]::new()
}

class SingleObjectSyncResult
{
    [ValidateSet("Create", "Update", "Delete", "Other")]
    [string] $Action = "Other"

    [string] $StartTime = [DateTime]::UtcNow.ToUniversalTime()

    [string] $EndTime = [string]::Empty

    [ProvisioningSystem] $SourceSystem = [ProvisioningSystem]::new()

    [ProvisioningSystem] $TargetSystem = [ProvisioningSystem]::new()

    [ProvisioningIdentity] $SourceIdentity = [ProvisioningIdentity]::new()

    [ProvisioningIdentity] $TargetIdentity = [ProvisioningIdentity]::new()

    [StatusInfo] $StatusInfo = [StatusInfo]::new()

    [System.Collections.Generic.List[ProvisioningProperty]] $ModifiedProperties = [System.Collections.Generic.List[ProvisioningProperty]]::new()

    [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps = [System.Collections.Generic.List[ProvisioningStep]]::new()
}

#-------------------------------------------------------------------------
#endregion class definitions
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
#region common helper functions
#-------------------------------------------------------------------------

function Write-Result
{
    param
    (
        [Parameter(Mandatory=$True)]
        [SingleObjectSyncResult] $Result,
        [Parameter(Mandatory=$True)]
        [bool] $HtmlReport
    )

    $Result.EndTime = [DateTime]::UtcNow.ToUniversalTime()

    if ($HtmlReport)
    {
        Write-HtmlReport -Result $Result
    }

    $ResultJson = ConvertTo-Json -InputObject $Result -Depth 3 -Compress
    Write-Output $ResultJson
}

function Test-StatusInfoFailure
{
    param
    (
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $IsFailure = $StatusInfo.Status -eq "Failure"
    Write-Output $IsFailure
}

# Trim whitespace from each RDN in DN
function Format-DistinguishedName
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName
    )

    # https://ldapwiki.com/wiki/Escaping%20Distinguished%20Name
    # Replace \, with \2C
    $Result = $DistinguishedName.Replace("\,", "\2C")

    # Split DN into RDN
    # Trim whitespace from each RDN
    $Result = $Result.Split(",") | ForEach-Object { $_.Trim() }

    # Join RDN's to DN
    $Result = $Result -join ","

    # Replace \2C back to \,
    $Result = $Result.Replace("\2C", "\,")

    Write-Output $Result
}

function Get-DomainComponent
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $Index = $DistinguishedName.IndexOf("DC=", [StringComparison]::OrdinalIgnoreCase)
    if ($Index -eq -1)
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Could not get the domain component from the distinguished name `"$DistinguishedName`""
        $StatusInfo.RecommendedAction = "The distinguished name input should have domain component. Domain component is a sequence of relative distinguished names (RDN) connected by commas where each RDN is in the form DC=value. For example: DC=contoso,DC=com"
        return
    }

    $DomainComponent = $DistinguishedName.Substring($Index)
    Write-Output $DomainComponent
}

function Get-StagingModeEnabled
{
    param
    (
        [Parameter(Mandatory=$True)]
        [bool] $StagingModePresent,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    if ($StagingModePresent)
    {
        $StagingModeEnabled = $True
    }
    else
    {
        try
        {
            $SchedulerSettings = Get-ADSyncScheduler
        }
        catch
        {
            $StatusInfo.Status = "Failure"
            $StatusInfo.Reason = "Exception thrown while running `"Get-ADSyncScheduler`" to determine if Staging Mode enabled"
            $StatusInfo.AdditionalDetails = $_.Exception.Message
            $StatusInfo.RecommendedAction = $Script:GenericRecommendedAction
            return
        }
        $StagingModeEnabled = $SchedulerSettings.StagingModeEnabled
    }

    Write-Output $StagingModeEnabled
}

function Initialize-ADConnectorAndPartition
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [ref] $ConnectorRef,
        [Parameter(Mandatory=$True)]
        [ref] $PartitionRef,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $DomainComponent = Get-DomainComponent -DistinguishedName $DistinguishedName -StatusInfo $StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $StatusInfo)
    {
        return
    }

    try
    {
        $ConnectorList = Get-ADSyncConnector
    }
    catch
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Exception thrown while running `"Get-ADSyncConnector`" to initialize the Active Directory connector and partition"
        $StatusInfo.AdditionalDetails = $_.Exception.Message
        $StatusInfo.RecommendedAction = $Script:GenericRecommendedAction
        return
    }

    foreach ($Connector in $ConnectorList)
    {
        foreach ($Partition in $Connector.Partitions)
        {
            if ($DomainComponent.Equals($Partition.DN, [StringComparison]::OrdinalIgnoreCase))
            {
                $ConnectorRef.Value = $Connector
                $PartitionRef.Value = $Partition
                return
            }
        }
    }

    $StatusInfo.Status = "Failure"
    $StatusInfo.Reason = "Could not find the connector for partition with the distinguished name `"$DomainComponent`""
    $StatusInfo.RecommendedAction = "Include the partition `"$DomainComponent`" in the list of domains that should be synced. To learn more on how to do this, please see: $Script:FilteringDomainLink"
}

function Initialize-AADConnectorAndPartition
{
    param
    (
        [Parameter(Mandatory=$True)]
        [ref] $ConnectorRef,
        [Parameter(Mandatory=$True)]
        [ref] $PartitionRef,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $AzureADConnectorGuid = [Microsoft.IdentityManagement.PowerShell.ObjectModel.Constants]::AzureADConnectorGuid

    $Connector = $Null
    $Partition = $Null

    try
    {
        $Connector = Get-ADSyncConnector -Identifier $AzureADConnectorGuid
    }
    catch
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.AdditionalDetails = $_.Exception.Message
        $StatusInfo.Reason = "Could not get the Azure Active Directory connector with the Guid: $AzureADConnectorGuid"
        $StatusInfo.RecommendedAction = "Please use the Azure AD Connect Wizard to configure your Azure Active Directory connector."
        return
    }

    if ($Connector.Partitions.Count -ne 1)
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Could not get the default partition for the Azure Active Directory connector"
        $StatusInfo.RecommendedAction = "Please use the Azure AD Connect Wizard to configure your Azure Active Directory connector."
        return
    }
    $Partition = $Connector.Partitions | Select-Object -First 1

    $ConnectorRef.Value = $Connector
    $PartitionRef.Value = $Partition
}

function Get-ConnectorSystemName
{
    param
    (
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector
    )

    $AzureADConnectorGuid = [Microsoft.IdentityManagement.PowerShell.ObjectModel.Constants]::AzureADConnectorGuid

    if ($Connector.Identifier -eq [guid]::new($AzureADConnectorGuid))
    {
        Write-Output "Azure Active Directory"
    }
    else
    {
        Write-Output "Active Directory"
    }
}

function Get-ADConnectorAccountName
{
    param
    (
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector
    )

    $SystemName = Get-ConnectorSystemName -Connector $Connector
    if ($SystemName -ne "Active Directory")
    {
        return
    }

    $ConnectorAccountName = "$($Connector.ConnectivityParameters['forest-login-domain'].Value)\$($Connector.ConnectivityParameters['forest-login-user'].Value)"
    Write-Output $ConnectorAccountName
}

# Get the most specific container / nearest ancestor to the object distinguished name from the container list.
function Get-MostSpecificContainer
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [System.Collections.Generic.List[string]] $ContainerList
    )

    $Length = 0
    $MostSpecificContainer = [string]::Empty
    foreach ($Container in $ContainerList)
    {
        if ($DistinguishedName.EndsWith($Container, [StringComparison]::OrdinalIgnoreCase))
        {
            if ($Container.Length -gt $Length)
            {
                $MostSpecificContainer = $Container
                $Length = $Container.Length
            }
        }
    }
    Write-Output $MostSpecificContainer
}

function Get-AdDirectoryObject
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$False)]
        [System.Collections.Generic.List[string]] $PropertiesToRetrieve = $Null,
        [Parameter(Mandatory=$False)]
        [StatusInfo] $StatusInfo
    )

    $SystemName = Get-ConnectorSystemName -Connector $Connector
    if ($SystemName -ne "Active Directory")
    {
        return
    }

    try
    {
        $AdObject = Search-ADSyncDirectoryObjects -AdConnectorId $Connector.Identifier -LdapFilter "(distinguishedName=$DistinguishedName)" -PropertiesToRetrieve $PropertiesToRetrieve -SearchScope Subtree -SizeLimit 1
        $AdObject = $AdObject | Select-Object -First 1
    }
    catch
    {
        if ($Null -ne $StatusInfo)
        {
            $StatusInfo.AdditionalDetails = $_.Exception.Message
        }
        $AdObject = $Null
    }

    if ($Null -ne $StatusInfo -and $Null -eq $AdObject)
    {
        $ConnectorAccountName = Get-ADConnectorAccountName -Connector $Connector
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Could not find an object in on-premises Active Directory with distinguished name `"$DistinguishedName`" using Connector account `"$ConnectorAccountName`" credentials."
        $StatusInfo.RecommendedAction = "Please ensure the object exists in on-premises Active Directory and the distinguished name is correct. If the object exists in Active Directory and the distinguished name is correct, verify the Connector account `"$ConnectorAccountName`" has sufficient permissions to read the object. To learn more on how to configure AD DS connector account permissions, please see: $Script:ConfigureAdDsConnectorAccountLink"
    }

    Write-Output $AdObject
}

function Get-CsObject
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$False)]
        [StatusInfo] $StatusInfo
    )

    try
    {
        $CsObject = Get-ADSyncCSObject -DistinguishedName $DistinguishedName -ConnectorIdentifier $Connector.Identifier
    }
    catch
    {
        if ($Null -ne $StatusInfo)
        {
            $StatusInfo.AdditionalDetails = $_.Exception.Message
        }
        $CsObject = $Null
    }

    if ($Null -ne $StatusInfo -and $Null -eq $CsObject)
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Could not find an object with distinguished name `"$DistinguishedName`" in `"$($Connector.Name)`" Connector Space."
        $StatusInfo.RecommendedAction = $Script:GenericRecommendedAction
    }

    Write-Output $CsObject
}

# Add Recommended Action, Reason, etc to Status Info using the Error Code
function Add-StatusInfoDetails
{
    param
    (
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    switch ($StatusInfo.ErrorCode)
    {
        "attributevaluemustbeunique"
        {
            $StatusInfo.RecommendedAction = "To learn how to fix AttributeValueMustBeUnique Export Error, please see: $Script:ExportErrorAttributeValueMustBeUniqueLink"
        }
        "datavalidationfailed"
        {
            $StatusInfo.RecommendedAction = "To learn how to fix DataValidationFailed Export Error, please see: $Script:ExportErrorDataValidationFailedLink"
        }
        "federateddomainchangeerror"
        {
            $StatusInfo.RecommendedAction = "To learn how to fix FederatedDomainChangeError Export Error, please see: $Script:ExportErrorFederatedDomainChangeErrorLink"
        }
        "invalidsoftmatch"
        {
            $StatusInfo.RecommendedAction = "To learn how to fix InvalidSoftMatch Export Error, please see: $Script:ExportErrorInvalidSoftMatchLink"
        }
        "largeobject"
        {
            $StatusInfo.RecommendedAction = "To learn how to fix LargeObject Export Error, please see: $Script:ExportErrorLargeObjectLink"
        }
        "no-start-ma-already-running"
        {
            $StatusInfo.Reason = "The run step failed to start because the run profile name specified in the RunProfileName parameter is already running. You can run only one run profile of a management agent at a time. However, you can run several run profiles at the same time if the run profiles are from different management agents."
            $StatusInfo.RecommendedAction = "Stop the run profile that is currently running or wait until the run profile has finished running before starting another management agent run profile. To learn more on how to disable the scheduler, please see: $Script:SchedulerDisableLink . To learn more on how to stop a synchronization cycle, please see: $Script:SchedulerStopLink"
        }
        "no-start-ma-update-in-progress"
        {
            $StatusInfo.Reason = "The run step failed to start because a new management agent is being created or an existing management agent is being modified or deleted. You cannot run a management agent when a new management agent is being created or an existing management agent is being modified or deleted."
            $StatusInfo.RecommendedAction = "Wait until the management agent has been created, deleted, or modified before starting a management agent run profile. To learn more on how to disable the scheduler, please see: $Script:SchedulerDisableLink . To learn more on how to stop a synchronization cycle, please see: $Script:SchedulerStopLink"
        }
        "objecttypemismatch"
        {
            $StatusInfo.RecommendedAction = "To learn how to fix ObjectTypeMismatch Export Error, please see: $Script:ExportErrorObjectTypeMismatchLink"
        }
    }
}

function Invoke-SpecificObjectRunProfileHelper
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [string] $RunProfileName,
        [Parameter(Mandatory=$True)]
        [int] $StepNumber,
        [Parameter(Mandatory=$False)]
        [StatusInfo] $StatusInfo
    )

    try
    {
        $RunProfileResult = Invoke-ADSyncSpecificObjectRunProfile -DistinguishedName $DistinguishedName -ConnectorIdentifier $Connector.Identifier -RunProfileName $RunProfileName -StepNumber $StepNumber
    }
    catch
    {
        if ($Null -ne $StatusInfo)
        {
            $StatusInfo.Status = "Failure"
            $StatusInfo.Reason = "Exception thrown while running Invoke-ADSyncSpecificObjectRunProfile -DistinguishedName `"$DistinguishedName`" -ConnectorIdentifier `"$($Connector.Identifier)`" -RunProfileName `"$RunProfileName`" -StepNumber $StepNumber"
            $StatusInfo.AdditionalDetails = $_.Exception.Message
            $StatusInfo.RecommendedAction = $Script:GenericRecommendedAction
        }
        return
    }

    $RunStepResult = $RunProfileResult.RunStepResults | Select-Object -First 1
    if ($Null -ne $StatusInfo -and $Null -eq $RunStepResult)
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.ErrorCode = $RunProfileResult.Result
        $StatusInfo.Reason = "RunStepResult in RunProfileResult is Null on Invoke-ADSyncSpecificObjectRunProfile -DistinguishedName `"$DistinguishedName`" -ConnectorIdentifier `"$($Connector.Identifier)`" -RunProfileName `"$RunProfileName`" -StepNumber $StepNumber"
        $StatusInfo.RecommendedAction = $Script:GenericRecommendedAction
        Add-StatusInfoDetails -StatusInfo $StatusInfo
    }

    Write-Output $RunStepResult
}

function Get-RunStepNumber
{
    param
    (
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.ConnectorPartition] $Partition,
        [Parameter(Mandatory=$True)]
        [string] $RunProfileName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask] $RunStepTask,
        [Parameter(Mandatory=$False)]
        [StatusInfo] $StatusInfo
    )

    $RunProfile = $Connector.RunProfiles | Where-Object { $RunProfileName.Equals($_.Name, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
    if ($Null -eq $RunProfile)
    {
        if ($Null -ne $StatusInfo)
        {
            $StatusInfo.Status = "Failure"
            $StatusInfo.Reason = "Connector `"$($Connector.Name)`" does not have Run Profile `"$RunProfileName`" configured"
            $StatusInfo.RecommendedAction = "Create Run Profile `"$RunProfileName`" for the connector `"$($Connector.Name)`" partition `"$($Partition.Name)`". To learn more on how to do this, please see: $Script:FilteringDomainLink"
        }
        return
    }

    $RunStepNumber = 0
    foreach ($RunStep in $RunProfile.RunSteps)
    {
        $RunStepNumber = $RunStepNumber + 1
        if ($RunStep.TaskType -ne $RunStepTask -or $RunStep.PartitionIdentifier -ne $Partition.Identifier)
        {
            continue
        }
        Write-Output $RunStepNumber
        return
    }

    if ($Null -ne $StatusInfo)
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Run Step `"$RunStepTask`" is not configured for connector `"$($Connector.Name)`" partition `"$($Partition.Name)`" for Run Profile `"$RunProfileName`""
        $StatusInfo.RecommendedAction = "Add Run Steps `"$RunStepTask`" to the Run Profile `"$RunProfileName`" to include the connector `"$($Connector.Name)`" partition `"$($Partition.Name)`". To learn more on how to do this, please see: $Script:FilteringDomainLink"
    }
}

function Add-RunStepResultStatusInfo
{
    param
    (
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepResult] $RunStepResult,
        [Parameter(Mandatory=$True)]
        [ValidateSet("import", "export")]
        [string] $StepType,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $ConnectionXmlDocument = [System.Xml.XmlDocument]::new()
    $ConnectionXmlDocument.LoadXml("<connection>$($RunStepResult.ConnectorConnectionInformationXml)</connection>")
    $ConnectionXmlElement = $ConnectionXmlDocument.SelectSingleNode("connection")
    $ConnectionResult = $ConnectionXmlElement."connection-result"
    if ($Null -ne $ConnectionResult -and $ConnectionResult -ne "success")
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.ErrorCode = $ConnectionResult
        $StatusInfo.AdditionalDetails = $ConnectionXmlElement.SelectSingleNode("connection-log/incident/cd-error")."error-literal"
        Add-StatusInfoDetails -StatusInfo $StatusInfo
        return
    }

    $DiscoveryErrorsXmlDocument = [System.Xml.XmlDocument]::new()
    $DiscoveryErrorsXmlDocument.LoadXml($RunStepResult.ConnectorDiscoveryErrors.ConnectorDiscoveryErrorsXml)
    $DiscoveryErrorsXmlElement = $DiscoveryErrorsXmlDocument.SelectSingleNode("ma-discovery-errors/ma-object-error")

    $SyncErrorsXmlDocument = [System.Xml.XmlDocument]::new()
    $SyncErrorsXmlDocument.LoadXml($RunStepResult.SyncErrors.SyncErrorsXml)
    $SyncErrorsXmlElement = $SyncErrorsXmlDocument.SelectSingleNode("synchronization-errors/$StepType-error")

    $ErrorsXmlElement = $Null
    if ($Null -ne $DiscoveryErrorsXmlElement)
    {
        $ErrorsXmlElement = $DiscoveryErrorsXmlElement
    }
    elseif ($Null -ne $SyncErrorsXmlElement)
    {
        $ErrorsXmlElement = $SyncErrorsXmlElement
    }

    if ($Null -ne $ErrorsXmlElement)
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.ErrorCode = $ErrorsXmlElement."error-type"
        if ($Null -ne $ErrorsXmlElement.SelectSingleNode("cd-error"))
        {
            $StatusInfo.Reason = $ErrorsXmlElement.SelectSingleNode("cd-error")."error-literal"
            if ($Null -ne $ErrorsXmlElement.SelectSingleNode("cd-error")."extra-error-details")
            {
                $StatusInfo.AdditionalDetails = $ErrorsXmlElement.SelectSingleNode("cd-error")."extra-error-details"
            }
            elseif ($Null -ne $ErrorsXmlElement.SelectSingleNode("cd-error")."server-error-detail")
            {
                $StatusInfo.AdditionalDetails = $ErrorsXmlElement.SelectSingleNode("cd-error")."server-error-detail"
            }
        }
        elseif ($Null -ne $ErrorsXmlElement.SelectSingleNode("change-not-reimported"))
        {
            $AttributeDeltaDictionary = Get-AttributeFragmentDictionary -FragmentType "delta" -FragmentXmlElement $ErrorsXmlElement.SelectSingleNode("change-not-reimported")."delta"
            $StatusInfo.AdditionalDetails = Get-AttributeInfoJson -AttributeDeltaDictionary $AttributeDeltaDictionary
        }
        Add-StatusInfoDetails -StatusInfo $StatusInfo
        return
    }

    if ($StepType -eq "import")
    {
        $StageCount = $RunStepResult.StageAdd + $RunStepResult.StageDelete + $RunStepResult.StageDeleteAdd + $RunStepResult.StageFailure + $RunStepResult.StageNoChange + $RunStepResult.StageRename + $RunStepResult.StageUpdate
        if ($StageCount -eq 0)
        {
            $StatusInfo.Status = "Failure"
            $StatusInfo.ErrorCode = "Filtered Object"
            return
        }
    }

    # Fail if Run Step Result is not success and not completed-transient-objects
    # Reason for completed-transient-objects: BUG 1185571 https://identitydivision.visualstudio.com/Engineering/_workitems/edit/1185571/
    if ($RunStepResult.StepResult -ne "success" -and $RunStepResult.StepResult -ne "completed-transient-objects")
    {
        $StatusInfo.Status = "Failure"
        $StatusInfo.ErrorCode = $RunStepResult.StepResult
        return
    }
}

function Add-SyncPreviewResultStatusInfo
{
    param
    (
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.PreviewResult] $SyncPreviewResult,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $SyncPreviewResultXmlDocument = [System.Xml.XmlDocument]::new()
    $SyncPreviewResultXmlDocument.LoadXml($SyncPreviewResult.SerializedXml)
    $ErrorXmlElement = $SyncPreviewResultXmlDocument.SelectSingleNode("preview/error")
    $DetailsXml = $SyncPreviewResultXmlDocument.SelectSingleNode("preview")."sync-generic-failure-error-details"

    $StatusInfo.Status = "Failure"
    $StatusInfo.ErrorCode = $ErrorXmlElement."type"
    $StatusInfo.Reason = $ErrorXmlElement."diagnosis"
    if (-not [string]::IsNullOrWhiteSpace($DetailsXml))
    {
        $DetailsXmlDocument = [System.Xml.XmlDocument]::new()
        $DetailsXmlDocument.LoadXml($DetailsXml)
        $StatusInfo.AdditionalDetails = $DetailsXmlDocument.SelectSingleNode("extension-error-info")."call-stack"
    }
    Add-StatusInfoDetails -StatusInfo $StatusInfo
}

function Add-CsObjectDetails
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [ValidateSet("pending-import", "unapplied-export")]
        [string] $Fragment,
        [Parameter(Mandatory=$False)]
        [System.Collections.Generic.Dictionary[[string], [AttributeInfo]]] $AttributeFlowDictionary,
        [Parameter(Mandatory=$True)]
        [System.Collections.Generic.Dictionary[[string], [string]]] $ProvisioningStepDetails,
        [Parameter(Mandatory=$False)]
        [StatusInfo] $StatusInfo
    )

    $CsObject = Get-CsObject -DistinguishedName $DistinguishedName -Connector $Connector -StatusInfo $StatusInfo
    if ($Null -eq $CsObject)
    {
        $ProvisioningStepDetails.Add("Object in `"$($Connector.Name)`" Connector Space", $False)
        return
    }
    $ProvisioningStepDetails.Add("Object in `"$($Connector.Name)`" Connector Space", $True)

    $CsObjectXml = [System.Xml.XmlDocument]::new()
    $CsObjectXml.LoadXml($CsObject.SerializedXml)

    $ObjectType = $CsObjectXml.SelectSingleNode("cs-objects/cs-object")."object-type"
    $ObjectOperation = $CsObjectXml.SelectSingleNode("cs-objects/cs-object/$Fragment/delta")."operation"

    $AttributeEntryDictionary = Get-AttributeFragmentDictionary -FragmentType "entry" -FragmentXmlElement $CsObjectXml.SelectSingleNode("cs-objects/cs-object/$Fragment-hologram/entry")
    $AttributeDeltaDictionary = Get-AttributeFragmentDictionary -FragmentType "delta" -FragmentXmlElement $CsObjectXml.SelectSingleNode("cs-objects/cs-object/$Fragment/delta")
    $AttributeInfoJson = Get-AttributeInfoJson -AttributeEntryDictionary $AttributeEntryDictionary -AttributeDeltaDictionary $AttributeDeltaDictionary -AttributeFlowDictionary $AttributeFlowDictionary

    $ProvisioningStepDetails.Add("Connector Space Object Type", $ObjectType)
    $ProvisioningStepDetails.Add("Connector Space Object Operation", $ObjectOperation)
    $ProvisioningStepDetails.Add("AttributeInfoJson", $AttributeInfoJson)

    Write-Output $CsObject
}

function Add-MvObjectDetails
{
    param
    (
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.PreviewResult] $SyncPreviewResult,
        [Parameter(Mandatory=$True)]
        [System.Collections.Generic.Dictionary[[string], [string]]] $ProvisioningStepDetails
    )

    $SyncPreviewResultXmlDocument = [System.Xml.XmlDocument]::new()
    $SyncPreviewResultXmlDocument.LoadXml($SyncPreviewResult.SerializedXml)

    $ObjectType = $SyncPreviewResultXmlDocument.SelectSingleNode("preview/mv/mv-object/entry/objectclass")."oc-value"
    $ObjectOperation = $SyncPreviewResultXmlDocument.SelectSingleNode("preview/mv/mv-changes/delta")."operation"

    $AttributeEntryDictionary = Get-AttributeFragmentDictionary -FragmentType "entry" -FragmentXmlElement $SyncPreviewResultXmlDocument.SelectSingleNode("preview/mv/mv-object/entry")
    $AttributeDeltaDictionary = Get-AttributeFragmentDictionary -FragmentType "delta" -FragmentXmlElement $SyncPreviewResultXmlDocument.SelectSingleNode("preview/mv/mv-changes/delta")
    $AttributeFlowDictionary = Get-AttributeFlowDictionary -AttributePrefix "mv" -AttributeFlowXmlElementList $SyncPreviewResultXmlDocument.SelectSingleNode("preview/import-flow-rules/import-attribute-flow")."import-flow"
    $AttributeInfoJson = Get-AttributeInfoJson -AttributeEntryDictionary $AttributeEntryDictionary -AttributeDeltaDictionary $AttributeDeltaDictionary -AttributeFlowDictionary $AttributeFlowDictionary

    $ProvisioningStepDetails.Add("Metaverse Object Type", $ObjectType)
    $ProvisioningStepDetails.Add("Metaverse Object Operation", $ObjectOperation)
    $ProvisioningStepDetails.Add("AttributeInfoJson", $AttributeInfoJson)
}

function Get-AttributeFragmentDictionary
{
    param
    (
        [Parameter(Mandatory=$True)]
        [ValidateSet("delta", "entry")]
        [string] $FragmentType,
        [Parameter(Mandatory=$True)]
        [System.Xml.XmlElement] $FragmentXmlElement
    )

    $AttrKeyValue = [System.Collections.Generic.Dictionary[[string], [string]]]::new()
    $AttrKeyValue.Add("attr", "value")
    $AttrKeyValue.Add("dn-attr", "dn-value")

    $AttributeFragmentDictionary = [System.Collections.Generic.Dictionary[[string], [AttributeInfo]]]::new()
    foreach ($AttrKey in $AttrKeyValue.Keys)
    {
        $AttrValue = $AttrKeyValue[$AttrKey]
        foreach ($AttributeXmlElement in $FragmentXmlElement.SelectNodes($AttrKey))
        {
            $AttributeInfo = [AttributeInfo]::new()

            # Attribute Name
            $AttributeName = $AttributeXmlElement."name"

            # Attribute IsMultiValued
            if ($Null -eq $AttributeXmlElement."multivalued" -or $AttributeXmlElement."multivalued" -eq "false")
            {
                $IsMultiValued = $False
            }
            else
            {
                $IsMultiValued = $True
            }
            $AttributeInfo.IsMultiValued = $IsMultiValued

            # Attribute Type
            if ($AttrKey -eq "dn-attr")
            {
                $AttributeType = "reference"
            }
            else
            {
                $AttributeType = $AttributeXmlElement."type"
            }
            $AttributeInfo.Type = $AttributeType

            if ($FragmentType -eq "delta")
            {
                # Attribute Operation
                if ($FragmentXmlElement."operation" -eq "add")
                {
                    $AttributeOperation = "add"
                }
                else
                {
                    $AttributeOperation = $AttributeXmlElement."operation"
                }
                $AttributeInfo.Operation = $AttributeOperation

                # Attribute Add & Delete
                $AttributeAdd = [System.Collections.Generic.List[string]]::new()
                $AttributeDelete = [System.Collections.Generic.List[string]]::new()
                $AttributeValueXmlElementList = $AttributeXmlElement.SelectNodes($AttrValue)
                foreach ($AttributeValueXml in $AttributeValueXmlElementList)
                {
                    if ($AttrKey -eq "dn-attr")
                    {
                        $Value = $AttributeValueXml."dn"
                    }
                    else
                    {
                        $Value = $AttributeValueXml
                    }
                    if ($Value.GetType() -eq [System.Xml.XmlElement])
                    {
                        $Value = $Value.InnerText
                    }
                    if ($Null -eq $AttributeValueXml."operation" -or $AttributeValueXml."operation" -eq "add")
                    {
                        $AttributeAdd.Add($Value)
                    }
                    else
                    {
                        $AttributeDelete.Add($Value)
                    }
                }
                $AttributeInfo.Add = $AttributeAdd -join ", "
                $AttributeInfo.Delete = $AttributeDelete -join ", "
            }
            else
            {
                # Attribute Value
                $AttributeValue = [System.Collections.Generic.List[string]]::new()
                $AttributeValueXmlElementList = $AttributeXmlElement.SelectNodes($AttrValue)
                foreach ($AttributeValueXml in $AttributeValueXmlElementList)
                {
                    if ($Null -ne $AttributeValueXml.Item("dn"))
                    {
                        $Value = $AttributeValueXml."dn"
                    }
                    else
                    {
                        $Value = $AttributeValueXml
                    }
                    if ($Value.GetType() -eq [System.Xml.XmlElement])
                    {
                        $Value = $Value.InnerText
                    }
                    $AttributeValue.Add($Value)
                }
                $AttributeInfo.Value = $AttributeValue -join ", "
            }

            $AttributeFragmentDictionary.Add($AttributeName, $AttributeInfo)
        }
    }
    Write-Output $AttributeFragmentDictionary
}

function Get-AttributeFlowDictionary
{
    param
    (
        [Parameter(Mandatory=$True)]
        [ValidateSet("mv", "cd")]
        [string] $AttributePrefix,
        [Parameter(Mandatory=$False)]
        [System.Collections.Generic.List[System.Xml.XmlElement]] $AttributeFlowXmlElementList
    )

    $AttributeFlowDictionary = [System.Collections.Generic.Dictionary[[string], [AttributeInfo]]]::new()
    foreach ($AttributeFlowXmlElement in $AttributeFlowXmlElementList)
    {
        $AttributeName = $AttributeFlowXmlElement."$($AttributePrefix)-attribute"
        $AttributeInfo = [AttributeInfo]::new()
        $AttributeInfo.SyncRule = $AttributeFlowXmlElement."status"
        $AttributeInfo.MappingType = $AttributeFlowXmlElement."mapping-type"
        $AttributeInfo.DataSource = $AttributeFlowXmlElement.SelectSingleNode("direct-mapping")."src-attribute"
        $AttributeFlowDictionary.Add($AttributeName, $AttributeInfo)
    }
    Write-Output $AttributeFlowDictionary
}

function Get-AttributeInfoJson
{
    param
    (
        [Parameter(Mandatory=$False)]
        [System.Collections.Generic.Dictionary[[string], [AttributeInfo]]] $AttributeEntryDictionary,
        [Parameter(Mandatory=$False)]
        [System.Collections.Generic.Dictionary[[string], [AttributeInfo]]] $AttributeDeltaDictionary,
        [Parameter(Mandatory=$False)]
        [System.Collections.Generic.Dictionary[[string], [AttributeInfo]]] $AttributeFlowDictionary
    )

    $AttributeInfoDictionary = [System.Collections.Generic.Dictionary[[string], [AttributeInfo]]]::new()

    if ($Null -ne $AttributeEntryDictionary)
    {
        foreach ($AttributeName in $AttributeEntryDictionary.Keys)
        {
            if (-not $AttributeInfoDictionary.ContainsKey($AttributeName))
            {
                $AttributeInfo = [AttributeInfo]::new()
                $AttributeInfo.Name = $AttributeName
                $AttributeInfoDictionary.Add($AttributeName, $AttributeInfo)
            }
            $AttributeInfoDictionary[$AttributeName].IsMultiValued = $AttributeEntryDictionary[$AttributeName].IsMultiValued
            $AttributeInfoDictionary[$AttributeName].Type = $AttributeEntryDictionary[$AttributeName].Type
            $AttributeInfoDictionary[$AttributeName].Value = $AttributeEntryDictionary[$AttributeName].Value
            $AttributeInfoDictionary[$AttributeName].Operation = "none"
        }
    }

    if ($Null -ne $AttributeDeltaDictionary)
    {
        foreach ($AttributeName in $AttributeDeltaDictionary.Keys)
        {
            if (-not $AttributeInfoDictionary.ContainsKey($AttributeName))
            {
                $AttributeInfo = [AttributeInfo]::new()
                $AttributeInfo.Name = $AttributeName
                $AttributeInfoDictionary.Add($AttributeName, $AttributeInfo)
            }
            $AttributeInfoDictionary[$AttributeName].IsMultiValued = $AttributeDeltaDictionary[$AttributeName].IsMultiValued
            $AttributeInfoDictionary[$AttributeName].Type = $AttributeDeltaDictionary[$AttributeName].Type
            $AttributeInfoDictionary[$AttributeName].Add = $AttributeDeltaDictionary[$AttributeName].Add
            $AttributeInfoDictionary[$AttributeName].Delete = $AttributeDeltaDictionary[$AttributeName].Delete
            $AttributeInfoDictionary[$AttributeName].Operation = $AttributeDeltaDictionary[$AttributeName].Operation
        }
    }

    if ($Null -ne $AttributeFlowDictionary)
    {
        foreach ($AttributeName in $AttributeFlowDictionary.Keys)
        {
            if (-not $AttributeInfoDictionary.ContainsKey($AttributeName))
            {
                $AttributeInfo = [AttributeInfo]::new()
                $AttributeInfo.Name = $AttributeName
                $AttributeInfoDictionary.Add($AttributeName, $AttributeInfo)
            }
            $AttributeInfoDictionary[$AttributeName].SyncRule = $AttributeFlowDictionary[$AttributeName].SyncRule
            $AttributeInfoDictionary[$AttributeName].MappingType = $AttributeFlowDictionary[$AttributeName].MappingType
            $AttributeInfoDictionary[$AttributeName].DataSource = $AttributeFlowDictionary[$AttributeName].DataSource
        }
    }

    $AttributeDetailsJson = ConvertTo-Json -InputObject $AttributeInfoDictionary.Values -Compress
    Write-Output $AttributeDetailsJson
}

function Get-TargetDistinguishedName
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $SourceDistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $SourceConnector,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $TargetConnector,
        [Parameter(Mandatory=$False)]
        [System.Collections.Generic.Dictionary[[string], [string]]] $ProvisioningStepDetails,
        [Parameter(Mandatory=$False)]
        [StatusInfo] $StatusInfo
    )

    $TargetDistinguishedName = [string]::Empty

    $SourceCsObject = Get-CsObject -DistinguishedName $SourceDistinguishedName -Connector $SourceConnector -StatusInfo $StatusInfo
    if ($Null -ne $ProvisioningStepDetails)
    {
        $ProvisioningStepDetails.Add("Source Object in `"$($SourceConnector.Name)`" Connector Space", ($Null -ne $SourceCsObject))
    }
    if ($Null -eq $SourceCsObject)
    {
        return
    }

    if ($Null -ne $SourceCsObject.ConnectedMVObjectId -and [guid]::Empty -ne $SourceCsObject.ConnectedMVObjectId)
    {
        try
        {
            $MvObject = Get-ADSyncMVObject -Identifier $SourceCsObject.ConnectedMVObjectId
        }
        catch
        {
            if ($Null -ne $StatusInfo)
            {
                $StatusInfo.AdditionalDetails = $_.Exception.Message
            }
            $MvObject = $Null
        }
    }
    else
    {
        $MvObject = $Null
    }
    if ($Null -ne $ProvisioningStepDetails)
    {
        $ProvisioningStepDetails.Add("Connected Object in Metaverse", ($Null -ne $MvObject))
    }
    if ($Null -eq $MvObject)
    {
        if ($Null -ne $StatusInfo)
        {
            $StatusInfo.Status = "Failure"
            $StatusInfo.Reason = "Could not find connected object in metaverse."
            $StatusInfo.RecommendedAction = "Please ensure the object is not being filtered by Attribute based Inbound filtering. To learn more about attribute based filtering, please see: $Script:FilteringAttributeLink . To learn more on how to customize a synchronization rule, please see: $Script:CustomSyncRuleLink"
        }
        return
    }

    # Get target distinguished name from the MV object lineage
    foreach ($Link in $MvObject.Lineage)
    {
        if ($TargetConnector.Identifier.Equals($Link.ConnectorId))
        {
            $TargetDistinguishedName = $Link.ConnectedCsObjectDN
            break
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($TargetDistinguishedName))
    {
        $TargetCsObject = Get-CsObject -DistinguishedName $TargetDistinguishedName -Connector $TargetConnector -StatusInfo $StatusInfo
    }
    else
    {
        $TargetCsObject = $Null
    }
    if ($Null -ne $ProvisioningStepDetails)
    {
        $ProvisioningStepDetails.Add("Connected Target Object in `"$($TargetConnector.Name)`" Connector Space", ($Null -ne $TargetCsObject))
    }
    if ($Null -eq $TargetCsObject)
    {
        if ($Null -ne $StatusInfo)
        {
            $StatusInfo.Status = "Failure"
            $StatusInfo.Reason = "Could not find connected object in `"$($TargetConnector.Name)`" Connector Space."
            $StatusInfo.RecommendedAction = "Please ensure the object is not being filtered by Attribute based Outbound filtering. To learn more about attribute based filtering, please see: $Script:FilteringAttributeLink . To learn more on how to customize a synchronization rule, please see: $Script:CustomSyncRuleLink"
        }
        return
    }

    Write-Output $TargetDistinguishedName
}

function Get-TargetAttributeFlowDictionary
{
    param
    (
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.PreviewResult] $SyncPreviewResult,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector
    )

    $TargetAttributeFlowDictionary = $Null

    $SyncPreviewResultXmlDocument = [System.Xml.XmlDocument]::new()
    $SyncPreviewResultXmlDocument.LoadXml($SyncPreviewResult.SerializedXml)
    $CsExportXmlElementList = $SyncPreviewResultXmlDocument.SelectSingleNode("preview")."cs-export"

    foreach ($CsExportXmlElement in $CsExportXmlElementList)
    {
        if ([guid]::new($CsExportXmlElement.SelectSingleNode("export-before-change")."ma-id") -eq $Connector.Identifier)
        {
            $TargetAttributeFlowDictionary = Get-AttributeFlowDictionary -AttributePrefix "cd" -AttributeFlowXmlElementList $CsExportXmlElement.SelectSingleNode("export-flow-rules/export-attribute-flow")."export-flow"
            break
        }
    }

    Write-Output $TargetAttributeFlowDictionary
}

function Get-OutOfScopeSyncRules
{
    param
    (
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.PreviewResult] $SyncPreviewResult
    )

    $OutOfScopeSyncRules = [System.Collections.Generic.List[string]]::new()
    foreach ($EntryModificationDiagnosticsData in $SyncPreviewResult.PreviewDiagnosticsData.EntryModificationDiagnosticsDataList)
    {
        foreach ($OutOfScopeSyncRule in $EntryModificationDiagnosticsData.ScopeModuleDiagnosticsData.OutOfScopeSyncRules)
        {
            $SyncRuleName = $OutOfScopeSyncRule.SyncRuleName
            if ($OutOfScopeSyncRule.SourceObjectMarkedForDeletion)
            {
                $OutOfScopeSyncRules.Add("$SyncRuleName (Source object marked for deletion)")
            }
            elseif ($OutOfScopeSyncRule.Disabled)
            {
                $OutOfScopeSyncRules.Add("$SyncRuleName (Sync rule disabled)")
            }
            else
            {
                $ScopeConditions = $OutOfScopeSyncRule.ScopeConditionGroups | ForEach-Object { "[$($_.Attribute) $($_.ComparisonOperator) $($_.ComparisonValue)]" }
                $ScopeConditions = $ScopeConditions -join ", "
                $OutOfScopeSyncRules.Add("$SyncRuleName (Scope conditions not satisfied: $ScopeConditions)")
            }
        }
    }
    Write-Output $OutOfScopeSyncRules
}

function Set-ProvisioningIdentity
{
    param
    (
        [Parameter(Mandatory=$True)]
        [AllowNull()]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject] $CsObject,
        [Parameter(Mandatory=$True)]
        [ProvisioningIdentity] $ProvisioningIdentity
    )

    if ($Null -eq $CsObject)
    {
        return
    }

    $ProvisioningIdentity.Type = $CsObject.ObjectType

    $AzureADConnectorGuid = [Microsoft.IdentityManagement.PowerShell.ObjectModel.Constants]::AzureADConnectorGuid
    if ($CsObject.ConnectorId -eq [guid]::new($AzureADConnectorGuid))
    {
        if ($CsObject.Attributes.Contains("cloudAnchor"))
        {
            $CloudAnchor = $CsObject.Attributes["cloudAnchor"].Values | Select-Object -First 1
            $ProvisioningIdentity.Id = [string]::new($CloudAnchor).Split("_")[1]
        }

        if ($CsObject.Attributes.Contains("userPrincipalName"))
        {
            $ProvisioningIdentity.Name = $CsObject.Attributes["userPrincipalName"].Values | Select-Object -First 1
        }
        elseif ($CsObject.Attributes.Contains("displayName"))
        {
            $ProvisioningIdentity.Name = $CsObject.Attributes["displayName"].Values | Select-Object -First 1
        }
    }
    else
    {
        $ProvisioningIdentity.Name = $CsObject.DistinguishedName

        if ($CsObject.Attributes.Contains("objectGUID"))
        {
            $Base64EncodedObjectGuid = $CsObject.Attributes["objectGUID"].Values | Select-Object -First 1
            $ObjectGuid = [guid]::new([System.Convert]::FromBase64String($Base64EncodedObjectGuid))
            $ProvisioningIdentity.Id = $ObjectGuid.ToString()
        }
    }
}

function Get-Action
{
    param
    (
        [Parameter(Mandatory=$False)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject] $CsObject
    )

    $Action = "Other"

    if ($Null -eq $CsObject)
    {
        Write-Output $Action
    }

    $CsObjectXml = [System.Xml.XmlDocument]::new()
    $CsObjectXml.LoadXml($CsObject.SerializedXml)
    $ObjectOperation = $CsObjectXml.SelectSingleNode("cs-objects/cs-object/unapplied-export/delta")."operation"

    switch ($ObjectOperation)
    {
        "add" { $Action = "Create" }
        "update" { $Action = "Update" }
        "delete" { $Action = "Delete" }
    }

    Write-Output $Action
}

function Compare-ListOfStrings
{
    param
    (
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]] $List1,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]] $List2
    )

    $CountMap = [System.Collections.Generic.Dictionary[[string], [int]]]::new()
    foreach ($Value in $List1)
    {
        if (-not $CountMap.ContainsKey($Value))
        {
            $CountMap.Add($Value, 0)
        }
        $CountMap[$Value] = $CountMap[$Value] + 1
    }
    foreach ($Value in $List2)
    {
        if (-not $CountMap.ContainsKey($Value))
        {
            $CountMap.Add($Value, 0)
        }
        $CountMap[$Value] = $CountMap[$Value] - 1
    }
    foreach ($Value in $CountMap.Values)
    {
        if ($Value -ne 0)
        {
            Write-Output $False
            return
        }
    }
    Write-Output $True
}

function Set-ProvisioningModifiedProperties
{
    param
    (
        [Parameter(Mandatory=$True)]
        [AllowNull()]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject] $CsObjectBeforeSync,
        [Parameter(Mandatory=$True)]
        [AllowNull()]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.CsObject] $CsObjectAfterSync,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningProperty]] $ProvisioningModifiedProperties
    )

    class CustomProvisioningProperty
    {
        [string] $Name

        [bool] $IsMultiValued

        [System.Collections.Generic.List[string]] $OldValues = [System.Collections.Generic.List[string]]::new()

        [System.Collections.Generic.List[string]] $NewValues = [System.Collections.Generic.List[string]]::new()
    }

    $PropertiesMap = [System.Collections.Generic.Dictionary[[string], [CustomProvisioningProperty]]]::new()

    if ($Null -ne $CsObjectBeforeSync)
    {
        foreach ($Attribute in $CsObjectBeforeSync.Attributes)
        {
            $AttributeName = $Attribute.Name
            if (-not $PropertiesMap.ContainsKey($AttributeName))
            {
                $PropertyObject = [CustomProvisioningProperty]::new()
                $PropertyObject.Name = $AttributeName
                $PropertyObject.IsMultiValued = $Attribute.IsMultiValued
                $PropertiesMap.Add($AttributeName, $PropertyObject)
            }
            $PropertyObject = $PropertiesMap[$AttributeName]
            $PropertyObject.OldValues = $Attribute.Values
        }
    }

    if ($Null -ne $CsObjectAfterSync)
    {
        foreach ($Attribute in $CsObjectAfterSync.Attributes)
        {
            $AttributeName = $Attribute.Name
            if (-not $PropertiesMap.ContainsKey($AttributeName))
            {
                $PropertyObject = [CustomProvisioningProperty]::new()
                $PropertyObject.Name = $AttributeName
                $PropertyObject.IsMultiValued = $Attribute.IsMultiValued
                $PropertiesMap.Add($AttributeName, $PropertyObject)
            }
            $PropertyObject = $PropertiesMap[$AttributeName]
            $PropertyObject.NewValues = $Attribute.Values
        }
    }

    foreach ($PropertyObject in $PropertiesMap.Values)
    {
        if ($PropertyObject.IsMultiValued)
        {
            $ListsAreEqual = Compare-ListOfStrings -List1 $PropertyObject.OldValues -List2 $PropertyObject.NewValues
            if (-not $ListsAreEqual)
            {
                $ProvisioningProperty = [ProvisioningProperty]::new()
                $ProvisioningProperty.Name = "$($PropertyObject.Name) - Count"
                $ProvisioningProperty.OldValue = $PropertyObject.OldValues.Count
                $ProvisioningProperty.NewValue = $PropertyObject.NewValues.Count
                $ProvisioningModifiedProperties.Add($ProvisioningProperty)
            }
        }
        else
        {
            $OldValue = $PropertyObject.OldValues | Select-Object -First 1
            $NewValue = $PropertyObject.NewValues | Select-Object -First 1
            if ($OldValue -ne $NewValue)
            {
                $ProvisioningProperty = [ProvisioningProperty]::new()
                $ProvisioningProperty.Name = $PropertyObject.Name
                $ProvisioningProperty.OldValue = $OldValue
                $ProvisioningProperty.NewValue = $NewValue
                $ProvisioningModifiedProperties.Add($ProvisioningProperty)
            }
        }
    }
}

#-------------------------------------------------------------------------
#endregion common helper functions
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
#region scoping step functions
#-------------------------------------------------------------------------

function Invoke-ProvisioningStepScopingDomain
{
    param
    (
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.ConnectorPartition] $Partition,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $ProvisioningStep = [ProvisioningStep]::new()
    $ProvisioningSteps.Add($ProvisioningStep)

    $ProvisioningStep.Type = "Scoping"
    $ProvisioningStep.Name = "ScopingDomain"
    $ProvisioningStep.Description = "Determine if Object's Domain (Connector:$($Connector.Name)) (Partition:$($Partition.Name)) in sync scope"

    if ($Partition.Selected)
    {
        $ProvisioningStep.Details.Add("Connector `"$($Connector.Name)`" Partition `"$($Partition.Name)`" selected", $True)
    }
    else
    {
        $ProvisioningStep.Details.Add("Connector `"$($Connector.Name)`" Partition `"$($Partition.Name)`" selected", $False)
        $ProvisioningStep.Status = "Failure"
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Object is not present in sync scope. Object belongs to connector `"$($Connector.Name)`" partition `"$($Partition.Name)`" which is not selected."
        $StatusInfo.RecommendedAction = "Include the connector `"$($Connector.Name)`" partition `"$($Partition.Name)`" in the list of domains that should be synced. To learn more on how to do this, please see: $Script:FilteringDomainLink"
        return
    }

    $ConnectorRunProfileSteps = [System.Collections.Generic.Dictionary[[string], [System.Collections.Generic.List[Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStep]]]]::new()
    foreach ($ConnectorRunProfile in $Connector.RunProfiles)
    {
        $ConnectorRunProfileSteps.Add($ConnectorRunProfile.Name, $ConnectorRunProfile.RunSteps)
    }

    $RunProfileStepTasks = [System.Collections.Generic.Dictionary[[string], [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]]]::new()
    $RunProfileStepTasks.Add("Full Import", [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]::FullImport)
    $RunProfileStepTasks.Add("Full Synchronization", [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]::FullSynchronization)
    $RunProfileStepTasks.Add("Delta Import", [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]::DeltaImport)
    $RunProfileStepTasks.Add("Delta Synchronization", [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]::DeltaSynchronization)
    $RunProfileStepTasks.Add("Export", [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]::Export)
    $RunProfileStepTasks.Add("Specific Object Import", [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]::SpecificObjectImport)
    $RunProfileStepTasks.Add("Specific Object Export", [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]::SpecificObjectExport)

    $RunProfilesMissingStep = [System.Collections.Generic.List[string]]::new()
    foreach ($RunProfileName in $RunProfileStepTasks.Keys)
    {
        $IsMissingRunStep = $True

        if ($ConnectorRunProfileSteps.ContainsKey($RunProfileName))
        {
            foreach ($RunStep in $ConnectorRunProfileSteps[$RunProfileName])
            {
                if ($RunStep.PartitionIdentifier -eq $Partition.Identifier -and $RunStep.TaskType -eq $RunProfileStepTasks[$RunProfileName])
                {
                    $ProvisioningStep.Details.Add("Connector `"$($Connector.Name)`" Run Profile `"$RunProfileName`" contains Run Step for Partition `"$($Partition.Name)`"", $True)
                    $IsMissingRunStep = $False
                    break
                }
            }
        }

        if ($IsMissingRunStep)
        {
            $ProvisioningStep.Details.Add("Connector `"$($Connector.Name)`" Run Profile `"$RunProfileName`" contains Run Step for Partition `"$($Partition.Name)`"", $False)
            $RunProfilesMissingStep.Add($RunProfileName)
        }
    }

    if ($RunProfilesMissingStep.Count -gt 0)
    {
        $RunProfilesMissingStepOutput = $RunProfilesMissingStep -join ", "
        $ProvisioningStep.Status = "Failure"
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "No run steps are configured for connector `"$($Connector.Name)`" partition `"$($Partition.Name)`" for run profile(s): $RunProfilesMissingStepOutput"
        $StatusInfo.RecommendedAction = "Add run steps to the run profile(s): $RunProfilesMissingStepOutput to include the connector `"$($Connector.Name)`" partition `"$($Partition.Name)`". To learn more on how to do this, please see: $Script:FilteringDomainLink"
        return
    }

    $ProvisioningStep.Status = "Success"
}

function Invoke-ProvisioningStepScopingOrganizationalUnit
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.ConnectorPartition] $Partition,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $ProvisioningStep = [ProvisioningStep]::new()
    $ProvisioningSteps.Add($ProvisioningStep)

    $ProvisioningStep.Type = "Scoping"
    $ProvisioningStep.Name = "ScopingOrganizationalUnit"
    $ProvisioningStep.Description = "Determine if Object's Organizational Unit in sync scope"

    $SystemName = Get-ConnectorSystemName -Connector $Connector
    if ($SystemName -ne "Active Directory")
    {
        $ProvisioningStep.Status = "Skipped"
        return
    }

    $ContainerInclusionList = $Partition.ConnectorPartitionScope.ContainerInclusionList
    $ContainerExclusionList = $Partition.ConnectorPartitionScope.ContainerExclusionList

    $ProvisioningStep.Details.Add("Partition `"$($Partition.Name)`" container inclusion list", $($ContainerInclusionList -join ", "))
    $ProvisioningStep.Details.Add("Partition `"$($Partition.Name)`" container exclusion list", $($ContainerExclusionList -join ", "))

    $MostSpecificContainerInInclusionList = Get-MostSpecificContainer -DistinguishedName $DistinguishedName -ContainerList $ContainerInclusionList
    $MostSpecificContainerInExclusionList = Get-MostSpecificContainer -DistinguishedName $DistinguishedName -ContainerList $ContainerExclusionList

    $IsExcluded = $False
    if (-not [string]::IsNullOrEmpty($MostSpecificContainerInInclusionList) -and -not [string]::IsNullOrEmpty($MostSpecificContainerInExclusionList))
    {
        if ($MostSpecificContainerInInclusionList.Length -lt $MostSpecificContainerInExclusionList.Length)
        {
            $IsExcluded = $True
        }
    }
    elseif (-not [string]::IsNullOrEmpty($MostSpecificContainerInExclusionList))
    {
        $IsExcluded = $True
    }

    if ($IsExcluded)
    {
        $ProvisioningStep.Status = "Failure"
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Object is not present in sync scope. Object belongs to container `"$MostSpecificContainerInExclusionList`" that is excluded from syncing."
        $StatusInfo.RecommendedAction = "Include the container `"$MostSpecificContainerInExclusionList`" in the list of organizational units that should be synced. To learn more on how to do this, please see: $Script:FilteringOrganizationalUnitLink"
        return
    }

    $ProvisioningStep.Status = "Success"
}

function Invoke-ProvisioningStepScopingConnectivity
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $ProvisioningStep = [ProvisioningStep]::new()
    $ProvisioningSteps.Add($ProvisioningStep)

    $ProvisioningStep.Type = "Scoping"
    $ProvisioningStep.Name = "ScopingConnectivity"
    $ProvisioningStep.Description = "Determine if Object is accessible using connector account credentials"

    $SystemName = Get-ConnectorSystemName -Connector $Connector
    if ($SystemName -ne "Active Directory")
    {
        $ProvisioningStep.Status = "Skipped"
        return
    }

    $ConnectorAccountName = Get-ADConnectorAccountName -Connector $Connector

    $AdObject = Get-AdDirectoryObject -DistinguishedName $DistinguishedName -Connector $Connector -PropertiesToRetrieve "distinguishedName" -StatusInfo $StatusInfo
    if ($Null -eq $AdObject)
    {
        $ProvisioningStep.Details.Add("Object accessible using connector account `"$ConnectorAccountName`" credentials", $False)
        $ProvisioningStep.Status = "Failure"
        return
    }

    $ProvisioningStep.Details.Add("Object accessible using connector account `"$ConnectorAccountName`" credentials", $True)
    $ProvisioningStep.Status = "Success"
}
function Invoke-ProvisioningStepScopingObjectType
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $ProvisioningStep = [ProvisioningStep]::new()
    $ProvisioningSteps.Add($ProvisioningStep)

    $ProvisioningStep.Type = "Scoping"
    $ProvisioningStep.Name = "ScopingObjectType"
    $ProvisioningStep.Description = "Determine if Object's Type in sync scope"

    $SystemName = Get-ConnectorSystemName -Connector $Connector
    if ($SystemName -ne "Active Directory")
    {
        $ProvisioningStep.Status = "Skipped"
        return
    }

    $ConnectorObjectInclusionList = $Connector.ObjectInclusionList
    $ProvisioningStep.Details.Add("Connector `"$($Connector.Name)`" object type inclusion list", $($ConnectorObjectInclusionList -join ", "))

    $AdObject = Get-AdDirectoryObject -DistinguishedName $DistinguishedName -Connector $Connector -PropertiesToRetrieve "objectClass" -StatusInfo $StatusInfo
    if ($Null -eq $AdObject)
    {
        $ProvisioningStep.Status = "Failure"
        return
    }

    $AdObjectClass = $AdObject["objectClass"]
    $AdObjectType = $AdObjectClass[$AdObjectClass.Count - 1]
    $ProvisioningStep.Details.Add("Object type in Active Directory", $AdObjectType)

    if ($ConnectorObjectInclusionList -notcontains $AdObjectType)
    {
        $ConnectorObjectInclusionListOutput = $ConnectorObjectInclusionList -join ", "
        $ProvisioningStep.Status = "Failure"
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Object is not present in sync scope. Object type `"$AdObjectType`" is not part of the connector object type inclusion list: $ConnectorObjectInclusionListOutput"
        $StatusInfo.RecommendedAction = "Please ensure the object type `"$AdObjectType`" is being used in a sync rule for this connector. To learn more on how to customize a synchronization rule, please see: $Script:CustomSyncRuleLink . If object type `"$AdObjectType`" is not available in the Synchronization Rules Editor, use the Wizard to refresh the directory schema. To learn more on how to refresh the directory schema, please see: $Script:RefreshDirectorySchemaLink"
        return
    }

    $ProvisioningStep.Status = "Success"
}
function Invoke-ProvisioningStepScopingGroup
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $ProvisioningStep = [ProvisioningStep]::new()
    $ProvisioningSteps.Add($ProvisioningStep)

    $ProvisioningStep.Type = "Scoping"
    $ProvisioningStep.Name = "ScopingGroup"
    $ProvisioningStep.Description = "Determine if Object is in sync scope if Group Filtering enabled"

    $SystemName = Get-ConnectorSystemName -Connector $Connector
    if ($SystemName -ne "Active Directory")
    {
        $ProvisioningStep.Status = "Skipped"
        return
    }

    try
    {
        $GlobalSettings = Get-ADSyncGlobalSettings
    }
    catch
    {
        $ProvisioningStep.Status = "Failure"
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Exception thrown while running `"Get-ADSyncGlobalSettings`" to determine if Group Filtering enabled"
        $StatusInfo.AdditionalDetails = $_.Exception.Message
        $StatusInfo.RecommendedAction = $Script:GenericRecommendedAction
        return
    }
    $GroupFilteringEnabled = $GlobalSettings.Parameters["Microsoft.OptionalFeature.GroupFiltering"].Value -eq "True"
    $GroupDistinguishedName = $Connector.GlobalParameters["Connector.GroupFilteringGroupDn"].Value

    if ($GroupFilteringEnabled -and -not [string]::IsNullOrWhiteSpace($GroupDistinguishedName))
    {
        $ProvisioningStep.Details.Add("Group filtering enabled", $True)
        $ProvisioningStep.Details.Add("Group distinguished name", $GroupDistinguishedName)

        $AdGroupObject = Get-AdDirectoryObject -DistinguishedName $GroupDistinguishedName -Connector $Connector -PropertiesToRetrieve "member" -StatusInfo $StatusInfo
        if ($Null -eq $AdGroupObject)
        {
            $ProvisioningStep.Status = "Failure"
            return
        }

        $AdGroupMembers = $AdGroupObject["member"]
        if ($AdGroupMembers -notcontains $DistinguishedName)
        {
            $ProvisioningStep.Details.Add("Object is member of Group `"$GroupDistinguishedName`" in Active Directory", $False)
            $ProvisioningStep.Status = "Failure"
            $StatusInfo.Status = "Failure"
            $StatusInfo.Reason = "Object is not present in sync scope. Group filtering is enabled and the object is not a member of filtering group `"$GroupDistinguishedName`" in the Active Directory"
            $StatusInfo.RecommendedAction = "Add the object to group `"$GroupDistinguishedName`" or change group filtering settings. To learn more on how to do this, please see: $Script:FilteringGroupLink"
            return
        }
        $ProvisioningStep.Details.Add("Object is member of Group `"$GroupDistinguishedName`" in Active Directory", $True)

        $CsGroupObject = Get-CsObject -DistinguishedName $GroupDistinguishedName -Connector $Connector
        if ($Null -ne $CsGroupObject)
        {
            $CsGroupMembers = $CsGroupObject.Attributes["member"].Values
        }
        else
        {
            $CsGroupMembers = $Null
        }

        if ($Null -eq $CsGroupObject -or $CsGroupMembers -notcontains $DistinguishedName)
        {
            $ProvisioningStep.Details.Add("Object is member of Group `"$GroupDistinguishedName`" in Active Directory Connector Space", $False)
            $ProvisioningStep.Status = "Failure"
            $StatusInfo.Status = "Failure"
            $StatusInfo.Reason = "Object is not present in sync scope. Group filtering is enabled and the object is a member of filtering group `"$GroupDistinguishedName`" in the Active Directory but is not a member of the group in the Active Directory Connector Space"
            $StatusInfo.RecommendedAction = "Please ensure an import operation has been run on the Active Directory Connector `"$($Connector.Name)`" since the object was added to group `"$GroupDistinguishedName`" and that there were no errors during the import. To learn more on how to use the sync service manager operations tab, please see: $Script:SyncServiceManagerUiOperationsLink"
            return
        }
        $ProvisioningStep.Details.Add("Object is member of Group `"$GroupDistinguishedName`" in Active Directory Connector Space", $True)
    }
    else
    {
        $ProvisioningStep.Details.Add("Group filtering enabled", $False)
    }

    $ProvisioningStep.Status = "Success"
}

#-------------------------------------------------------------------------
#endregion scoping step functions
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
#region import step functions
#-------------------------------------------------------------------------

function Invoke-ProvisioningStepImportSpecificObject
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.ConnectorPartition] $Partition,
        [Parameter(Mandatory=$False)]
        [System.Collections.Generic.Dictionary[[string], [AttributeInfo]]] $AttributeFlowDictionary,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps,
        [Parameter(Mandatory=$False)]
        [StatusInfo] $StatusInfo
    )

    $ProvisioningStep = [ProvisioningStep]::new()
    $ProvisioningSteps.Add($ProvisioningStep)

    $ProvisioningStep.Type = "Import"
    $ProvisioningStep.Name = "ImportSpecificObject"
    $SystemName = Get-ConnectorSystemName -Connector $Connector
    $ProvisioningStep.Description = "Import Object from $SystemName to `"$($Connector.Name)`" Connector Space"

    $RunProfileName = [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunProfileName]::SpecificObjectImport
    $RunStepTask = [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]::SpecificObjectImport
    $StepNumber = Get-RunStepNumber -Connector $Connector -Partition $Partition -RunProfileName $RunProfileName -RunStepTask $RunStepTask -StatusInfo $StatusInfo
    if ($Null -eq $StepNumber)
    {
        $ProvisioningStep.Status = "Failure"
        return
    }

    $RunStepResult = Invoke-SpecificObjectRunProfileHelper -DistinguishedName $DistinguishedName -Connector $Connector -RunProfileName $RunProfileName -StepNumber $StepNumber -StatusInfo $StatusInfo
    if ($Null -eq $RunStepResult)
    {
        $ProvisioningStep.Status = "Failure"
        return
    }
    $ProvisioningStep.Details.Add("Run step result", $RunStepResult.StepResult)

    if ($Null -ne $StatusInfo)
    {
        Add-RunStepResultStatusInfo -RunStepResult $RunStepResult -StepType "import" -StatusInfo $StatusInfo
        if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
        {
            $ProvisioningStep.Status = "Failure"
            return
        }
    }

    $CsObject = Add-CsObjectDetails -DistinguishedName $DistinguishedName -Connector $Connector -Fragment "pending-import" -AttributeFlowDictionary $AttributeFlowDictionary -ProvisioningStepDetails $ProvisioningStep.Details -StatusInfo $StatusInfo
    if ($Null -eq $CsObject)
    {
        $ProvisioningStep.Status = "Failure"
        return
    }
    Write-Output $CsObject

    $ProvisioningStep.Status = "Success"
}

#-------------------------------------------------------------------------
#endregion import step functions
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
#region export step functions
#-------------------------------------------------------------------------

function Invoke-ProvisioningStepExportSpecificObject
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $Connector,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.ConnectorPartition] $Partition,
        [Parameter(Mandatory=$False)]
        [System.Collections.Generic.Dictionary[[string], [AttributeInfo]]] $AttributeFlowDictionary,
        [Parameter(Mandatory=$False)]
        [bool] $StagingModeEnabled,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps,
        [Parameter(Mandatory=$False)]
        [StatusInfo] $StatusInfo
    )

    $ProvisioningStep = [ProvisioningStep]::new()
    $ProvisioningSteps.Add($ProvisioningStep)

    $ProvisioningStep.Type = "Export"
    $ProvisioningStep.Name = "ExportSpecificObject"
    $SystemName = Get-ConnectorSystemName -Connector $Connector
    $ProvisioningStep.Description = "Export Object from `"$($Connector.Name)`" Connector Space to $SystemName"

    $RunProfileName = [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunProfileName]::SpecificObjectExport
    $RunStepTask = [Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepTask]::SpecificObjectExport
    $StepNumber = Get-RunStepNumber -Connector $Connector -Partition $Partition -RunProfileName $RunProfileName -RunStepTask $RunStepTask -StatusInfo $StatusInfo
    if ($Null -eq $StepNumber)
    {
        $ProvisioningStep.Status = "Failure"
        return
    }

    $CsObject = Add-CsObjectDetails -DistinguishedName $DistinguishedName -Connector $Connector -Fragment "unapplied-export" -AttributeFlowDictionary $AttributeFlowDictionary -ProvisioningStepDetails $ProvisioningStep.Details -StatusInfo $StatusInfo
    if ($Null -eq $CsObject)
    {
        $ProvisioningStep.Status = "Failure"
        return
    }
    Write-Output $CsObject

    if ($Null -ne $StagingModeEnabled -and $StagingModeEnabled)
    {
        $ProvisioningStep.Description = "[Staging Mode] Export Object from `"$($Connector.Name)`" Connector Space to $SystemName"
        $ProvisioningStep.Status = "Success"
        return
    }

    $RunStepResult = Invoke-SpecificObjectRunProfileHelper -DistinguishedName $DistinguishedName -Connector $Connector -RunProfileName $RunProfileName -StepNumber $StepNumber -StatusInfo $StatusInfo
    if ($Null -eq $RunStepResult)
    {
        $ProvisioningStep.Status = "Failure"
        return
    }
    $ProvisioningStep.Details.Add("Run step result", $RunStepResult.StepResult)

    if ($Null -ne $StatusInfo)
    {
        Add-RunStepResultStatusInfo -RunStepResult $RunStepResult -StepType "export" -StatusInfo $StatusInfo
        if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
        {
            $ProvisioningStep.Status = "Failure"
            return
        }
    }

    $ProvisioningStep.Status = "Success"
}

#-------------------------------------------------------------------------
#endregion export step functions
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
#region sync step functions
#-------------------------------------------------------------------------

function Invoke-ProvisioningStepSyncSpecificObject
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $SourceDistinguishedName,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $SourceConnector,
        [Parameter(Mandatory=$True)]
        [Microsoft.IdentityManagement.PowerShell.ObjectModel.Connector] $TargetConnector,
        [Parameter(Mandatory=$True)]
        [ref] $ActionRef,
        [Parameter(Mandatory=$True)]
        [ref] $TargetDistinguishedNameRef,
        [Parameter(Mandatory=$True)]
        [ref] $TargetAttributeFlowDictionaryRef,
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps,
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    $ProvisioningStep = [ProvisioningStep]::new()
    $ProvisioningSteps.Add($ProvisioningStep)

    $ProvisioningStep.Type = "Sync"
    $ProvisioningStep.Name = "SyncSpecificObject"
    $ProvisioningStep.Description = "Sync Object from `"$($SourceConnector.Name)`" Connector Space"

    $TargetDistinguishedName = [string]::Empty
    $TargetAttributeFlowDictionary = $Null

    try
    {
        $SyncPreviewResult = Sync-ADSyncCsObject -DistinguishedName $SourceDistinguishedName -ConnectorIdentifier $SourceConnector.Identifier -Commit
    }
    catch
    {
        $ProvisioningStep.Status = "Failure"
        $StatusInfo.Status = "Failure"
        $StatusInfo.Reason = "Exception thrown while running `"Sync-ADSyncCsObject`""
        $StatusInfo.AdditionalDetails = $_.Exception.Message
        $StatusInfo.RecommendedAction = $Script:GenericRecommendedAction
        return
    }

    if ($Null -ne $SyncPreviewResult.ErrorXml)
    {
        $ProvisioningStep.Details.Add("Sync step result", "error")
        $ProvisioningStep.Status = "Failure"
        Add-SyncPreviewResultStatusInfo -SyncPreviewResult $SyncPreviewResult -StatusInfo $StatusInfo
        return
    }
    $ProvisioningStep.Details.Add("Sync step result", "success")

    $OutOfScopeSyncRules = Get-OutOfScopeSyncRules -SyncPreviewResult $SyncPreviewResult
    $ProvisioningStep.Details.Add("Out of scope sync rules", $($OutOfScopeSyncRules -join ", "))

    Add-MvObjectDetails -SyncPreviewResult $SyncPreviewResult -ProvisioningStepDetails $ProvisioningStep.Details

    $TargetAttributeFlowDictionary = Get-TargetAttributeFlowDictionary -SyncPreviewResult $SyncPreviewResult -Connector $TargetConnector

    $TargetDistinguishedName = Get-TargetDistinguishedName -SourceDistinguishedName $SourceDistinguishedName -SourceConnector $SourceConnector -TargetConnector $TargetConnector -ProvisioningStepDetails $ProvisioningStep.Details -StatusInfo $StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        $ProvisioningStep.Status = "Failure"
        return
    }
    $TargetCsObject = Get-CsObject -DistinguishedName $TargetDistinguishedName -Connector $TargetConnector
    $ActionRef.Value = Get-Action -CsObject $TargetCsObject

    $TargetDistinguishedNameRef.Value = $TargetDistinguishedName
    $TargetAttributeFlowDictionaryRef.Value = $TargetAttributeFlowDictionary

    $ProvisioningStep.Status = "Success"
}

#-------------------------------------------------------------------------
#endregion sync step functions
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
#region html report functions
#-------------------------------------------------------------------------

function New-HtmlElement
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $ElementName,
        [Parameter(Mandatory=$False)]
        [Hashtable] $Attributes,
        [Parameter(Mandatory=$False)]
        [string] $InnerHtml
    )

    $HtmlElement = "<$ElementName"

    if ($Null -ne $Attributes)
    {
        foreach ($AttributeName in $Attributes.Keys)
        {
            $HtmlElement += " $AttributeName"
            $AttributeValue = $Attributes[$AttributeName]
            if ($Null -ne $AttributeValue)
            {
                $HtmlElement += "=`"$AttributeValue`""
            }
        }
    }

    if ([string]::IsNullOrEmpty($InnerHtml))
    {
        $HtmlElement += " />"
    }
    else
    {
        $HtmlElement += ">$InnerHtml</$ElementName>"
    }

    Write-Output $HtmlElement
}

function Get-IconChevron
{
    $Path = New-HtmlElement -ElementName "path" -Attributes @{"fill-rule"="evenodd"; d="M7.646 4.646a.5.5 0 0 1 .708 0l6 6a.5.5 0 0 1-.708.708L8 5.707l-5.646 5.647a.5.5 0 0 1-.708-.708l6-6z"}
    $Svg = New-HtmlElement -ElementName "svg" -Attributes @{width="1em"; height="1em"; viewBox="0 0 16 16"; class="bi bi-chevron-up"; fill="currentColor"; xmlns="http://www.w3.org/2000/svg"} -InnerHtml $Path
    Write-Output $Svg
}

function Get-IconCheck
{
    $Path = New-HtmlElement -ElementName "path" -Attributes @{"fill-rule"="evenodd"; d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zm-3.97-3.03a.75.75 0 0 0-1.08.022L7.477 9.417 5.384 7.323a.75.75 0 0 0-1.06 1.06L6.97 11.03a.75.75 0 0 0 1.079-.02l3.992-4.99a.75.75 0 0 0-.01-1.05z"}
    $Svg = New-HtmlElement -ElementName "svg" -Attributes @{width="1.25em"; height="1.25em"; viewBox="0 0 16 16"; class="bi bi-check-circle-fill"; fill="green"; xmlns="http://www.w3.org/2000/svg"} -InnerHtml $Path
    Write-Output $Svg
}

function Get-IconAlert
{
    $Path = New-HtmlElement -ElementName "path" -Attributes @{"fill-rule"="evenodd"; d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zM8 4a.905.905 0 0 0-.9.995l.35 3.507a.552.552 0 0 0 1.1 0l.35-3.507A.905.905 0 0 0 8 4zm.002 6a1 1 0 1 0 0 2 1 1 0 0 0 0-2z"}
    $Svg = New-HtmlElement -ElementName "svg" -Attributes @{width="1.25em"; height="1.25em"; viewBox="0 0 16 16"; class="bi bi-exclamation-circle-fill"; fill="red"; xmlns="http://www.w3.org/2000/svg"} -InnerHtml $Path
    Write-Output $Svg
}

function Get-HtmlHead
{
    $Title = New-HtmlElement -ElementName "title" -InnerHtml "ADSync Single Object Sync Result"
    $Icon = New-HtmlElement -ElementName "link" -Attributes @{rel="icon"; type="image/x-icon"; href="https://portal.azure.com//favicon.ico"}
    $CssBootstrap = New-HtmlElement -ElementName "link" -Attributes @{rel="stylesheet"; href="https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/css/bootstrap.min.css"; integrity="sha384-TX8t27EcRE3e/ihU7zmQxVncDAy5uIKz4rEkgIXeMed4M0jlfIDPvg6uqKI2xXr2"; crossorigin="anonymous"}
    $Style = New-HtmlElement -ElementName "style" -InnerHtml "[data-toggle=`"collapse`"].collapsed svg.bi.bi-chevron-up {transform: rotate(180deg);}"

    $Head = New-HtmlElement -ElementName "head" -InnerHtml "$Title $Icon $CssBootstrap $Style"
    Write-Output $Head
}

function Get-HtmlNavTabs
{
    $TabAttributes = @{class="nav-item"}

    $StepsLink = New-HtmlElement -ElementName "a" -Attributes @{class="nav-link active"; href="#steps"; "data-toggle"="tab"; "data-target"="#steps"} -InnerHtml "Steps"
    $StepsTab = New-HtmlElement -ElementName "li" -Attributes $TabAttributes -InnerHtml $StepsLink

    $TroubleshootingLink = New-HtmlElement -ElementName "a" -Attributes @{class="nav-link"; href="#troubleshooting"; "data-toggle"="tab"; "data-target"="#troubleshooting"} -InnerHtml "Troubleshooting & Recommendations"
    $TroubleshootingTab = New-HtmlElement -ElementName "li" -Attributes $TabAttributes -InnerHtml $TroubleshootingLink

    $ModifiedPropertiesLink = New-HtmlElement -ElementName "a" -Attributes @{class="nav-link"; href="#modifiedproperties"; "data-toggle"="tab"; "data-target"="#modifiedproperties"} -InnerHtml "Modified Properties"
    $ModifiedPropertiesTab = New-HtmlElement -ElementName "li" -Attributes $TabAttributes -InnerHtml $ModifiedPropertiesLink

    $SummaryLink = New-HtmlElement -ElementName "a" -Attributes @{class="nav-link"; href="#summary"; "data-toggle"="tab"; "data-target"="#summary"} -InnerHtml "Summary"
    $SummaryTab = New-HtmlElement -ElementName "li" -Attributes $TabAttributes -InnerHtml $SummaryLink

    $NavTabs = New-HtmlElement -ElementName "ul" -Attributes @{class="nav nav-tabs"} -InnerHtml "$StepsTab $TroubleshootingTab $ModifiedPropertiesTab $SummaryTab"
    Write-Output $NavTabs
}

function Get-HtmlAttributeInfoTable
{
    param
    (
        [Parameter(Mandatory=$True)]
        [System.Collections.Generic.List[AttributeInfo]] $AttributeInfoList,
        [Parameter(Mandatory=$True)]
        [int] $StepNumber
    )

    $SortedAttributeInfo = [System.Collections.Generic.List[AttributeInfo]]::new()
    $SortedAttributeInfo += $AttributeInfoList | Where-Object { $_.Operation -eq "add" }
    $SortedAttributeInfo += $AttributeInfoList | Where-Object { $_.Operation -eq "delete" }
    $SortedAttributeInfo += $AttributeInfoList | Where-Object { $_.Operation -eq "update" }
    $SortedAttributeInfo += $AttributeInfoList | Where-Object { $_.Operation -eq "none" }
    $SortedAttributeInfo += $AttributeInfoList | Where-Object { $_.Operation -ne "add" -and $_.Operation -ne "delete" -and $_.Operation -ne "update" -and $_.Operation -ne "none" }

    $AttributeInfoRows = New-HtmlElement -ElementName "tr" -InnerHtml $(New-HtmlElement -ElementName "th" -Attributes @{colspan="3"} -InnerHtml "Attribute Info")
    $AttributeInfoRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th") $(New-HtmlElement -ElementName "th" -InnerHtml "Name") $(New-HtmlElement -ElementName "th" -InnerHtml "Operation")"
    foreach ($Attribute in $SortedAttributeInfo)
    {
        $AttributeIndex = [Regex]::Replace($Attribute.Name, "[^\w]", "_")
        $AttributeInfoRow1 = "$(New-HtmlElement -ElementName "td" -Attributes @{width="1em"} -InnerHtml $(Get-IconChevron)) $(New-HtmlElement -ElementName "td" -InnerHtml $(New-HtmlElement -ElementName "a" -Attributes @{href="#steps-step$StepNumber-$AttributeIndex"} -InnerHtml $Attribute.Name)) $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.Operation)"
        $AttributeInfoRows += New-HtmlElement -ElementName "tr" -Attributes @{class="collapsed"; "data-toggle"="collapse"; "data-target"="#steps-step$StepNumber-$AttributeIndex"} -InnerHtml $AttributeInfoRow1

        $AttributeRow = New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Name") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.Name)"
        $AttributeRow += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Is Multi-Valued") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.IsMultiValued)"
        $AttributeRow += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Type") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.Type)"
        if (-not [string]::IsNullOrWhiteSpace($Attribute.Value))
        {
            $AttributeRow += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Value") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.Value)"
        }
        if (-not [string]::IsNullOrWhiteSpace($Attribute.Add))
        {
            $AttributeRow += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Value Add") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.Add)"
        }
        if (-not [string]::IsNullOrWhiteSpace($Attribute.Delete))
        {
            $AttributeRow += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Value Delete") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.Delete)"
        }
        $AttributeRow += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Operation") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.Operation)"
        if (-not [string]::IsNullOrWhiteSpace($Attribute.SyncRule))
        {
            $AttributeRow += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Sync Rule") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.SyncRule)"
        }
        if (-not [string]::IsNullOrWhiteSpace($Attribute.MappingType))
        {
            $AttributeRow += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Mapping Type") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.MappingType)"
        }
        if (-not [string]::IsNullOrWhiteSpace($Attribute.DataSource))
        {
            $AttributeRow += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Data Source") $(New-HtmlElement -ElementName "td" -InnerHtml $Attribute.DataSource)"
        }
        $AttributeTable = New-HtmlElement "table" -Attributes @{class="table table-sm table-borderless"} -InnerHtml $AttributeRow

        $AttributeInfoRow2 = "$(New-HtmlElement -ElementName "td") $(New-HtmlElement -ElementName "td"-Attributes @{colspan="2"} -InnerHtml $AttributeTable)"
        $AttributeInfoRows += New-HtmlElement -ElementName "tr" -Attributes @{id="steps-step$StepNumber-$AttributeIndex"; class="collapse table-secondary"} -InnerHtml $AttributeInfoRow2
    }

    $AttributeInfoTable = New-HtmlElement "table" -Attributes @{class="table table-sm"} -InnerHtml $AttributeInfoRows
    Write-Output $AttributeInfoTable
}

function Get-HtmlCardComponent
{
    param
    (
        [Parameter(Mandatory=$True)]
        [ProvisioningStep] $ProvisioningStep,
        [Parameter(Mandatory=$True)]
        [int] $StepNumber
    )

    if ($ProvisioningStep.Status -eq "success")
    {
        $Icon = Get-IconCheck
    }
    else
    {
        $Icon = Get-IconAlert
    }
    $ButtonText = "$(Get-IconChevron) $StepNumber. $($ProvisioningStep.Description) $Icon"
    $HeaderButton = New-HtmlElement -ElementName "button" -Attributes @{class="btn btn-link btn-block text-left collapsed"; type="button"; "data-toggle"="collapse"; "data-target"="#steps-step$StepNumber"} -InnerHtml $ButtonText
    $CardHeader = New-HtmlElement -ElementName "div" -Attributes @{class="card-header"} -InnerHtml $HeaderButton

    $ProvisioningStepRows = New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Status") $(New-HtmlElement -ElementName "td" -InnerHtml $ProvisioningStep.Status)"
    $ProvisioningStepRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Type") $(New-HtmlElement -ElementName "td" -InnerHtml $ProvisioningStep.Type)"
    $ProvisioningStepRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Name") $(New-HtmlElement -ElementName "td" -InnerHtml $ProvisioningStep.Name)"
    $ProvisioningStepRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Description") $(New-HtmlElement -ElementName "td" -InnerHtml $ProvisioningStep.Description)"
    $ProvisioningStepRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Date Time") $(New-HtmlElement -ElementName "td" -InnerHtml $ProvisioningStep.Timestamp)"
    $ProvisioningStepTable = New-HtmlElement "table" -Attributes @{class="table"} -InnerHtml $ProvisioningStepRows

    $ProvisioningStepDetailsTable = [string]::Empty
    if ($Null -ne $ProvisioningStep.Details -and $ProvisioningStep.Details.Count -gt 0)
    {
        $ProvisioningStepDetailsRows = New-HtmlElement -ElementName "tr" -InnerHtml $(New-HtmlElement -ElementName "th" -Attributes @{colspan="2"} -InnerHtml "Details")
        foreach ($DetailKey in $ProvisioningStep.Details.Keys)
        {
            if ($DetailKey -ne "AttributeInfoJson")
            {
                $DetailValue = $ProvisioningStep.Details[$DetailKey]
                $ProvisioningStepDetailsRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "td" -Attributes @{width="50%"} -InnerHtml $DetailKey) $(New-HtmlElement -ElementName "td" -Attributes @{width="50%"} -InnerHtml $DetailValue)"
            }
        }
        $ProvisioningStepDetailsTable = New-HtmlElement "table" -Attributes @{class="table"} -InnerHtml $ProvisioningStepDetailsRows
    }

    $AttributeInfoTable = [string]::Empty
    if ($Null -ne $ProvisioningStep.Details -and $ProvisioningStep.Details.ContainsKey("AttributeInfoJson"))
    {
        [System.Collections.Generic.List[AttributeInfo]] $AttributeInfoList = ConvertFrom-Json -InputObject $ProvisioningStep.Details["AttributeInfoJson"]
        if ($AttributeInfoList.Count -gt 0)
        {
            $AttributeInfoTable = Get-HtmlAttributeInfoTable -AttributeInfoList $AttributeInfoList -StepNumber $StepNumber
        }
    }

    $ProvisioningStepTable = New-HtmlElement "table" -Attributes @{class="table"} -InnerHtml $ProvisioningStepRows
    $CardBody = New-HtmlElement -ElementName "div" -Attributes @{id="steps-step$StepNumber"; class="collapse"; "data-parent"="#stepsaccordion"} -InnerHtml $(New-HtmlElement -ElementName "div" -Attributes @{class="card-body"} -InnerHtml "$ProvisioningStepTable $ProvisioningStepDetailsTable $AttributeInfoTable")

    $CardComponent = New-HtmlElement -ElementName "div" -Attributes @{class="card"} -InnerHtml "$CardHeader $CardBody"
    Write-Output $CardComponent
}

function Get-HtmlTabContentSteps
{
    param
    (
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps
    )

    if ($ProvisioningSteps.Count -eq 0)
    {
        $TabContentStepsRows =  New-HtmlElement -ElementName "tr" -InnerHtml $(New-HtmlElement -ElementName "th" -InnerHtml "There are no provisioning steps.")
        $TabContentSteps = New-HtmlElement -ElementName "table" -Attributes @{class="table"} -InnerHtml $TabContentStepsRows
    }
    else
    {
        $TabContentStepCards = [string]::Empty
        for ($Index = 0; $Index -lt $ProvisioningSteps.Count; $Index++)
        {
            $TabContentStepCards += Get-HtmlCardComponent -ProvisioningStep $ProvisioningSteps[$Index] -StepNumber $($Index + 1)
        }
        $TabContentSteps = New-HtmlElement -ElementName "div" -Attributes @{class="accordion"; id="stepsaccordion"} -InnerHtml $TabContentStepCards
    }

    Write-Output $TabContentSteps
}

function Format-HtmlTextLink
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $Text
    )

    # Replace words that starts with https:// with HTML hyperlink tag
    $Pattern = '(https?:\/\/[^\s]+)'
    $Substitution = '<a href="$1" target="_blank">$1</a>'

    $FormattedText = [regex]::Replace($Text, $Pattern, $Substitution)

    Write-Output $FormattedText
}

function Get-HtmlTabContentTroubleshooting
{
    param
    (
        [Parameter(Mandatory=$True)]
        [StatusInfo] $StatusInfo
    )

    if ($StatusInfo.Status -eq "success")
    {
        $TabContentTroubleshootingRows =  New-HtmlElement -ElementName "tr" -InnerHtml $(New-HtmlElement -ElementName "th" -InnerHtml "Looking good! There's nothing to troubleshoot.")
    }
    else
    {
        # Status
        $TabContentTroubleshootingRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Status") $(New-HtmlElement -ElementName "td" -InnerHtml $StatusInfo.Status)"

        if (-not [string]::IsNullOrWhiteSpace($StatusInfo.ErrorCode))
        {
            # Error Code
            $TabContentTroubleshootingRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Error Code") $(New-HtmlElement -ElementName "td" -InnerHtml $StatusInfo.ErrorCode)"
        }

        if (-not [string]::IsNullOrWhiteSpace($StatusInfo.Reason))
        {
            # Reason
            $Reason = Format-HtmlTextLink -Text $StatusInfo.Reason
            $TabContentTroubleshootingRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Reason") $(New-HtmlElement -ElementName "td" -InnerHtml $Reason)"
        }

        if (-not [string]::IsNullOrWhiteSpace($StatusInfo.AdditionalDetails))
        {
            # Additional Details
            $AdditionalDetails = Format-HtmlTextLink -Text $StatusInfo.AdditionalDetails
            $TabContentTroubleshootingRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Additional Details") $(New-HtmlElement -ElementName "td" -InnerHtml $AdditionalDetails)"
        }

        if (-not [string]::IsNullOrWhiteSpace($StatusInfo.ErrorCategory))
        {
            # Error Category
            $TabContentTroubleshootingRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Error Category") $(New-HtmlElement -ElementName "td" -InnerHtml $StatusInfo.ErrorCategory)"
        }

        if (-not [string]::IsNullOrWhiteSpace($StatusInfo.RecommendedAction))
        {
            # Recommended Action
            $RecommendedAction = Format-HtmlTextLink -Text $StatusInfo.RecommendedAction
            $TabContentTroubleshootingRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Recommended Action") $(New-HtmlElement -ElementName "td" -InnerHtml $RecommendedAction)"
        }
    }

    $TabContentTroubleshooting = New-HtmlElement "table" -Attributes @{class="table"} -InnerHtml $TabContentTroubleshootingRows
    Write-Output $TabContentTroubleshooting
}

function Get-ExportStepNumber
{
    param
    (
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningStep]] $ProvisioningSteps
    )

    if ($ProvisioningSteps.Count -gt 0)
    {
        for ($Index = 0; $Index -lt $ProvisioningSteps.Count; $Index++)
        {
            if ($ProvisioningSteps[$Index].Name -eq "ExportSpecificObject")
            {
                $ExportStepNumber = $Index + 1
                Write-Output $ExportStepNumber
            }
        }
    }
}

function Get-HtmlTabContentModifiedProperties
{
    param
    (
        [Parameter(Mandatory=$True)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[ProvisioningProperty]] $ModifiedProperties,
        [Parameter(Mandatory=$False)]
        [int] $StepNumber
    )

    if ($ModifiedProperties.Count -eq 0)
    {
        $TabContentModifiedPropertiesRows =  New-HtmlElement -ElementName "tr" -InnerHtml $(New-HtmlElement -ElementName "th" -InnerHtml "There are no modified properties.")
    }
    else
    {
        $TabContentModifiedPropertiesRows = [string]::Empty
        $TabContentModifiedPropertiesRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Name") $(New-HtmlElement -ElementName "th" -InnerHtml "Old Value") $(New-HtmlElement -ElementName "th" -InnerHtml "New Value")"
        foreach ($ModifiedProperty in $ModifiedProperties)
        {
            if ($Null -eq $StepNumber)
            {
                $PropertyName = $ModifiedProperty.Name
            }
            else
            {
                $Suffix = " - Count"
                $AttributeName = $ModifiedProperty.Name
                if ($AttributeName.EndsWith($Suffix, [StringComparison]::OrdinalIgnoreCase))
                {
                    $AttributeName = $AttributeName.Substring(0, $($AttributeName.Length - $Suffix.Length))
                }
                $AttributeIndex = [Regex]::Replace($AttributeName, "[^\w]", "_")
                $AttributeIdentifier = "steps-step$StepNumber-$AttributeIndex"
                $Script = "`$('[data-target=\'#steps\']').tab('show');`$('#steps-step$StepNumber').collapse('show');`$('#$AttributeIdentifier').collapse('show');`$('#$AttributeIdentifier').get(0).scrollIntoView();"
                $PropertyName = New-HtmlElement -ElementName "a" -Attributes @{href="#$AttributeIdentifier"; onclick=$Script} -InnerHtml $ModifiedProperty.Name
            }
            $TabContentModifiedPropertiesRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "td" -InnerHtml $PropertyName) $(New-HtmlElement -ElementName "td" -InnerHtml $ModifiedProperty.OldValue) $(New-HtmlElement -ElementName "td" -InnerHtml $ModifiedProperty.NewValue)"
        }
    }

    $TabContentModifiedProperties = New-HtmlElement -ElementName "table" -Attributes @{class="table"} -InnerHtml $TabContentModifiedPropertiesRows
    Write-Output $TabContentModifiedProperties
}

function Get-HtmlTabContentSummary
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string] $Action,
        [Parameter(Mandatory=$True)]
        [string] $StartTime,
        [Parameter(Mandatory=$True)]
        [string] $EndTime,
        [Parameter(Mandatory=$True)]
        [ProvisioningSystem] $SourceSystem,
        [Parameter(Mandatory=$True)]
        [ProvisioningSystem] $TargetSystem,
        [Parameter(Mandatory=$True)]
        [ProvisioningIdentity] $SourceIdentity,
        [Parameter(Mandatory=$True)]
        [ProvisioningIdentity] $TargetIdentity
    )

    $TabContentSummaryRows = [string]::Empty

    # Action
    $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Action") $(New-HtmlElement -ElementName "td" -InnerHtml $Action)"

    # Start Time
    $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Start Time") $(New-HtmlElement -ElementName "td" -InnerHtml $StartTime)"

    # End Time
    $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "End Time") $(New-HtmlElement -ElementName "td" -InnerHtml $EndTime)"

    # Source System
    $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Source System") $(New-HtmlElement -ElementName "td" -InnerHtml $SourceSystem.Name)"

    if (-not [string]::IsNullOrWhiteSpace($SourceIdentity.Name))
    {
        # Source Identity Name
        $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Source Object Name") $(New-HtmlElement -ElementName "td" -InnerHtml $SourceIdentity.Name)"
    }

    if (-not [string]::IsNullOrWhiteSpace($SourceIdentity.Type))
    {
        # Source Identity Type
        $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Source Object Type") $(New-HtmlElement -ElementName "td" -InnerHtml $SourceIdentity.Type)"
    }

    if (-not [string]::IsNullOrWhiteSpace($SourceIdentity.Id))
    {
        # Source Identity ID
        $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Source Object ID") $(New-HtmlElement -ElementName "td" -InnerHtml $SourceIdentity.Id)"
    }

    # Target System
    $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Target System") $(New-HtmlElement -ElementName "td" -InnerHtml $TargetSystem.Name)"

    if (-not [string]::IsNullOrWhiteSpace($TargetIdentity.Name))
    {
        # Target Identity Name
        $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Target Object Name") $(New-HtmlElement -ElementName "td" -InnerHtml $TargetIdentity.Name)"
    }

    if (-not [string]::IsNullOrWhiteSpace($TargetIdentity.Type))
    {
        # Target Identity Type
        $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Target Object Type") $(New-HtmlElement -ElementName "td" -InnerHtml $TargetIdentity.Type)"
    }

    if (-not [string]::IsNullOrWhiteSpace($TargetIdentity.Id))
    {
        # Target Identity ID
        $TabContentSummaryRows += New-HtmlElement -ElementName "tr" -InnerHtml "$(New-HtmlElement -ElementName "th" -InnerHtml "Target Object ID") $(New-HtmlElement -ElementName "td" -InnerHtml $TargetIdentity.Id)"
    }

    $TabContentSummary = New-HtmlElement -ElementName "table" -Attributes @{class="table"} -InnerHtml $TabContentSummaryRows
    Write-Output $TabContentSummary
}

function Get-HtmlTabContent
{
    param
    (
        [Parameter(Mandatory=$True)]
        [SingleObjectSyncResult] $Result
    )

    $StepsHtml = Get-HtmlTabContentSteps -ProvisioningSteps $Result.ProvisioningSteps
    $StepsContent = New-HtmlElement -ElementName "li" -Attributes @{class="tab-pane active"; id="steps"} -InnerHtml $StepsHtml

    $TroubleshootingHtml = Get-HtmlTabContentTroubleshooting -StatusInfo $Result.StatusInfo
    $TroubleshootingContent = New-HtmlElement -ElementName "li" -Attributes @{class="tab-pane"; id="troubleshooting"} -InnerHtml $TroubleshootingHtml

    $ExportStepNumber = Get-ExportStepNumber -ProvisioningSteps $Result.ProvisioningSteps
    $ModifiedPropertiesHtml = Get-HtmlTabContentModifiedProperties -ModifiedProperties $Result.ModifiedProperties -StepNumber $ExportStepNumber
    $ModifiedPropertiesContent = New-HtmlElement -ElementName "li" -Attributes @{class="tab-pane"; id="modifiedproperties"} -InnerHtml $ModifiedPropertiesHtml

    $SummaryHtml =  Get-HtmlTabContentSummary -Action $Result.Action -StartTime $Result.StartTime -EndTime $Result.EndTime -SourceSystem $Result.SourceSystem -TargetSystem $Result.TargetSystem -SourceIdentity $Result.SourceIdentity -TargetIdentity $Result.TargetIdentity
    $SummaryContent = New-HtmlElement -ElementName "div" -Attribute @{class="tab-pane"; id="summary"} -InnerHtml $SummaryHtml

    $TabContent = New-HtmlElement -ElementName "div" -Attributes @{class="tab-content"} -InnerHtml "$StepsContent $TroubleshootingContent $ModifiedPropertiesContent $SummaryContent"
    Write-Output $TabContent
}

function Get-HtmlMainContainer
{
    param
    (
        [Parameter(Mandatory=$True)]
        [SingleObjectSyncResult] $Result
    )

    $NavTabs = Get-HtmlNavTabs
    $TabContent = Get-HtmlTabContent -Result $Result
    $Main = New-HtmlElement -ElementName "main" -InnerHtml "$NavTabs $TabContent"
    $Container = New-HtmlElement -ElementName "div" -Attributes @{class="container"} -InnerHtml $Main
    Write-Output $Container
}

function Get-HtmlBody
{
    param
    (
        [Parameter(Mandatory=$True)]
        [SingleObjectSyncResult] $Result
    )

    $Brand = New-HtmlElement -ElementName "a" -Attributes @{class="navbar-brand"; href="#"} -InnerHtml "Microsoft Azure AD Connect: Single Object Sync Report"
    $Banner = New-HtmlElement -ElementName "nav" -Attributes @{class="navbar navbar-dark bg-dark"} -InnerHtml $Brand

    $MainContainer = Get-HtmlMainContainer -Result $Result

    $JsJquery = New-HtmlElement -ElementName "script" -Attributes @{src="https://code.jquery.com/jquery-3.5.1.slim.min.js"; integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj"; crossorigin="anonymous"} -InnerHtml " "
    $JsBootstrap = New-HtmlElement -ElementName "script" -Attributes @{src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/js/bootstrap.bundle.min.js"; integrity="sha384-ho+j7jyWK8fNQe+A12Hb8AhRq26LrZ/JpcUGGOn+Y7RsweNrtN/tE3MoK7ZeZDyx"; crossorigin="anonymous"} -InnerHtml " "

    $Body = New-HtmlElement -ElementName "body" -InnerHtml "$Banner $MainContainer $JsJquery $JsBootstrap"
    Write-Output $Body
}

function Write-HtmlReport
{
    param
    (
        [Parameter(Mandatory=$True)]
        [SingleObjectSyncResult] $Result
    )

    $Head = Get-HtmlHead
    $Body = Get-HtmlBody -Result $Result
    $ResultHtml = New-HtmlElement -ElementName "html" -InnerHtml "$Head $Body"

    $Path = "$Env:ProgramData\AADConnect\ADSyncObjectDiagnostics"
    if (-not (Test-Path -Path $Path))
    {
        New-Item -ItemType Directory -Path $Path -Force
    }
    $FileName = "ADSyncSingleObjectSyncResult-$(Get-Date -Format yyyyMMddHHmmss).htm"
    $FilePath = "$Path\$FileName"
    Out-File -FilePath $FilePath -InputObject $ResultHtml
    Start-Process -FilePath $FilePath
}

#-------------------------------------------------------------------------
#endregion html report functions
#-------------------------------------------------------------------------

#-------------------------------------------------------------------------
#region public functions to export
#-------------------------------------------------------------------------

function Invoke-ADSyncSingleObjectSync
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string] $DistinguishedName,
        [Parameter(Mandatory=$False)]
        [Switch] $StagingMode,
        [Parameter(Mandatory=$False)]
        [Switch] $NoHtmlReport
    )

    $HtmlReport = $True
    if ($NoHtmlReport.IsPresent)
    {
        $HtmlReport = $False
    }

    # Progress variables
    $ActivityName = "Single object sync in progress"
    $StepIndex = 0
    $TotalSteps = 14
    Write-Progress -Activity $ActivityName -Status "Running single object sync" -PercentComplete $($StepIndex++ / $TotalSteps * 100)

    # Source variables
    $SourceDistinguishedName = Format-DistinguishedName -DistinguishedName $DistinguishedName
    $SourceConnector = $Null
    $SourcePartition = $Null
    $SourceCsObject = $Null

    # Target variables
    $TargetDistinguishedName = [string]::Empty
    $TargetConnector = $Null
    $TargetPartition = $Null
    $TargetAttributeFlowDictionary = $Null
    $TargetCsObjectBeforeSync = $Null
    $TargetCsObjectAfterSync = $Null

    # Initialize Single Object Sync Result
    $Result = [SingleObjectSyncResult]::new()
    $Result.SourceIdentity.Name = $SourceDistinguishedName

    $StagingModeEnabled = Get-StagingModeEnabled -StagingModePresent $StagingMode.IsPresent -StatusInfo $Result.StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    Write-Progress -Activity $ActivityName -Status "Initializing Active Directory connector and partition" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    Initialize-ADConnectorAndPartition -DistinguishedName $SourceDistinguishedName -ConnectorRef ([ref]$SourceConnector) -PartitionRef ([ref]$SourcePartition) -StatusInfo $Result.StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    $Result.SourceSystem.Name = Get-ConnectorSystemName -Connector $SourceConnector

    Write-Progress -Activity $ActivityName -Status "Initializing Azure Active Directory connector and partition" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    Initialize-AADConnectorAndPartition -ConnectorRef ([ref]$TargetConnector) -PartitionRef ([ref]$TargetPartition) -StatusInfo $Result.StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    $Result.TargetSystem.Name = Get-ConnectorSystemName -Connector $TargetConnector

    Write-Progress -Activity $ActivityName -Status "Running provisioning step scoping Active Directory domain" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    Invoke-ProvisioningStepScopingDomain -Connector $SourceConnector -Partition $SourcePartition -ProvisioningSteps $Result.ProvisioningSteps -StatusInfo $Result.StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    Write-Progress -Activity $ActivityName -Status "Running provisioning step scoping Azure Active Directory domain" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    Invoke-ProvisioningStepScopingDomain -Connector $TargetConnector -Partition $TargetPartition -ProvisioningSteps $Result.ProvisioningSteps -StatusInfo $Result.StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    Write-Progress -Activity $ActivityName -Status "Running provisioning step scoping Active Directory organizational unit" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    Invoke-ProvisioningStepScopingOrganizationalUnit -DistinguishedName $SourceDistinguishedName -Connector $SourceConnector -Partition $SourcePartition -ProvisioningSteps $Result.ProvisioningSteps -StatusInfo $Result.StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    Write-Progress -Activity $ActivityName -Status "Running provisioning step scoping Active Directory connectivity" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    Invoke-ProvisioningStepScopingConnectivity -DistinguishedName $SourceDistinguishedName -Connector $SourceConnector -ProvisioningSteps $Result.ProvisioningSteps -StatusInfo $Result.StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    Write-Progress -Activity $ActivityName -Status "Running provisioning step scoping Active Directory object type" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    Invoke-ProvisioningStepScopingObjectType -DistinguishedName $SourceDistinguishedName -Connector $SourceConnector -ProvisioningSteps $Result.ProvisioningSteps -StatusInfo $Result.StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    Write-Progress -Activity $ActivityName -Status "Running provisioning step scoping Active Directory group filtering" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    Invoke-ProvisioningStepScopingGroup -DistinguishedName $SourceDistinguishedName -Connector $SourceConnector -ProvisioningSteps $Result.ProvisioningSteps -StatusInfo $Result.StatusInfo
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    Write-Progress -Activity $ActivityName -Status "Running provisioning step import from Active Directory" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    $SourceCsObject = Invoke-ProvisioningStepImportSpecificObject -DistinguishedName $SourceDistinguishedName -Connector $SourceConnector -Partition $SourcePartition -ProvisioningSteps $Result.ProvisioningSteps -StatusInfo $Result.StatusInfo
    Set-ProvisioningIdentity -CsObject $SourceCsObject -ProvisioningIdentity $Result.SourceIdentity
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    $TargetDistinguishedName = Get-TargetDistinguishedName -SourceDistinguishedName $SourceDistinguishedName -SourceConnector $SourceConnector -TargetConnector $TargetConnector

    Write-Progress -Activity $ActivityName -Status "Running provisioning step import from Azure Active Directory" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    if (-not [string]::IsNullOrWhiteSpace($TargetDistinguishedName))
    {
        $TargetCsObjectBeforeSync = Invoke-ProvisioningStepImportSpecificObject -DistinguishedName $TargetDistinguishedName -Connector $TargetConnector -Partition $TargetPartition -ProvisioningStep $Result.ProvisioningSteps
        Set-ProvisioningIdentity -CsObject $TargetCsObjectBeforeSync -ProvisioningIdentity $Result.TargetIdentity
    }

    Write-Progress -Activity $ActivityName -Status "Running provisioning step sync from Active Directory" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    $Action = $Result.Action
    Invoke-ProvisioningStepSyncSpecificObject -SourceDistinguishedName $SourceDistinguishedName -SourceConnector $SourceConnector -TargetConnector $TargetConnector -ActionRef ([ref]$Action) -TargetDistinguishedNameRef ([ref]$TargetDistinguishedName) -TargetAttributeFlowDictionaryRef ([ref]$TargetAttributeFlowDictionary) -ProvisioningSteps $Result.ProvisioningSteps -StatusInfo $Result.StatusInfo
    $Result.Action = $Action
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    Write-Progress -Activity $ActivityName -Status "Running provisioning step export to Azure Active Directory" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    $TargetCsObjectAfterSync = Invoke-ProvisioningStepExportSpecificObject -DistinguishedName $TargetDistinguishedName -Connector $TargetConnector -Partition $TargetPartition -AttributeFlowDictionary $TargetAttributeFlowDictionary -StagingModeEnabled $StagingModeEnabled -ProvisioningSteps $Result.ProvisioningSteps -StatusInfo $Result.StatusInfo
    Set-ProvisioningIdentity -CsObject $TargetCsObjectAfterSync -ProvisioningIdentity $Result.TargetIdentity
    if (Test-StatusInfoFailure -StatusInfo $Result.StatusInfo)
    {
        Write-Result -Result $Result -HtmlReport $HtmlReport
        return
    }

    Write-Progress -Activity $ActivityName -Status "Generating single object sync result" -PercentComplete $($StepIndex++ / $TotalSteps * 100)
    Set-ProvisioningModifiedProperties -CsObjectBeforeSync $TargetCsObjectBeforeSync -CsObjectAfterSync $TargetCsObjectAfterSync -ProvisioningModifiedProperties $Result.ModifiedProperties

    $Result.StatusInfo.Status = "Success"
    Write-Result -Result $Result -HtmlReport $HtmlReport
}

#-------------------------------------------------------------------------
#endregion public functions to export
#-------------------------------------------------------------------------

# SIG # Begin signature block
# MIIoOQYJKoZIhvcNAQcCoIIoKjCCKCYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAlnSivQ+aSE62s
# uXCNU0HEvtHPRx9a49YVQ4Mjv+KTMqCCDYIwggYAMIID6KADAgECAhMzAAADXJXz
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIHusQ4Vw
# bv+fHa5H20E3y9AI3yirtM8tlsjHw0IlVEwIMEIGCisGAQQBgjcCAQwxNDAyoBSA
# EgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20w
# DQYJKoZIhvcNAQEBBQAEggEANJTE+jrkoGoECcMpOvl8uBY3A75e3LtUUtxX9+xs
# N8yj8HD1g6lxirAK6rvGh4BiE1737C0nf3QtBbwHvN5LImi5vF26tX3yFGHtGTfD
# VqgDtWtvzMqFnrSGF5tm91+6aj6wWajk1MZOh/r212bT65uThtArk8Pew/dWlDzA
# fHDHckemo0evVhmKAfjzuvrUax4zv+mZwEpsk5zwp07k7HlYRqSU2G1MgOcw3Qa+
# VGRIUeMKFCamUTo/LnrNZAEOIn75dRmqjYdBEjcYxs8vvbl+/F0HGaN34J+q/vN8
# ZgIdC21eNZhdNnCUDMeExHryugDaP0lL5RjxJTjZ2Y2SEKGCF5cwgheTBgorBgEE
# AYI3AwMBMYIXgzCCF38GCSqGSIb3DQEHAqCCF3AwghdsAgEDMQ8wDQYJYIZIAWUD
# BAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoD
# ATAxMA0GCWCGSAFlAwQCAQUABCDDvx3H1PlzRdY48EOUrj2QRQ4OGecdGyWQm8z7
# yUdjwgIGZQQEMn4aGBMyMDIzMTAwNDE5MjgyMi45NDZaMASAAgH0oIHRpIHOMIHL
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxN
# aWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRT
# UyBFU046REMwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFNlcnZpY2WgghHtMIIHIDCCBQigAwIBAgITMwAAAdIhJDFKWL8tEQABAAAB
# 0jANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAe
# Fw0yMzA1MjUxOTEyMjFaFw0yNDAyMDExOTEyMjFaMIHLMQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmlj
# YSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046REMwMC0wNUUw
# LUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDcYIhC0QI/SPaT5+nYSBsSdhBP
# O2SXM40Vyyg8Fq1TPrMNDzxChxWUD7fbKwYGSsONgtjjVed5HSh5il75jNacb6Tr
# ZwuX+Q2++f2/8CCyu8TY0rxEInD3Tj52bWz5QRWVQejfdCA/n6ZzinhcZZ7+VelW
# gTfYC7rDrhX3TBX89elqXmISOVIWeXiRK8h9hH6SXgjhQGGQbf2bSM7uGkKzJ/pZ
# 2LvlTzq+mOW9iP2jcYEA4bpPeurpglLVUSnGGQLmjQp7Sdy1wE52WjPKdLnBF6Jb
# mSREM/Dj9Z7okxRNUjYSdgyvZ1LWSilhV/wegYXVQ6P9MKjRnE8CI5KMHmq7EsHh
# IBK0B99dFQydL1vduC7eWEjzz55Z/DyH6Hl2SPOf5KZ4lHf6MUwtgaf+MeZxkW0i
# xh/vL1mX8VsJTHa8AH+0l/9dnWzFMFFJFG7g95nHJ6MmYPrfmoeKORoyEQRsSus2
# qCrpMjg/P3Z9WJAtFGoXYMD19NrzG4UFPpVbl3N1XvG4/uldo1+anBpDYhxQU7k1
# gfHn6QxdUU0TsrJ/JCvLffS89b4VXlIaxnVF6QZh+J7xLUNGtEmj6dwPzoCfL7zq
# DZJvmsvYNk1lcbyVxMIgDFPoA2fZPXHF7dxahM2ZG7AAt3vZEiMtC6E/ciLRcIwz
# lJrBiHEenIPvxW15qwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFCC2n7cnR3ToP/kb
# EZ2XJFFmZ1kkMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1Ud
# HwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3Js
# L01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggr
# BgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIw
# MTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgw
# DgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCw5iq0Ey0LlAdz2Pcq
# chRwW5d+fitNISCvqD0E6W/AyiTk+TM3WhYTaxQ2pP6Or4qOV+Du7/L+k18gYr1p
# hshxVMVnXNcdjecMtTWUOVAwbJoeWHaAgknNIMzXK3+zguG5TVcLEh/CVMy1J7KP
# E8Q0Cz56NgWzd9urG+shSDKkKdhOYPXF970Mr1GCFFpe1oXjEy6aS+Heavp2wmy6
# 5mbu0AcUOPEn+hYqijgLXSPqvuFmOOo5UnSV66Dv5FdkqK7q5DReox9RPEZcHUa+
# 2BUKPjp+dQ3D4c9IH8727KjMD8OXZomD9A8Mr/fcDn5FI7lfZc8ghYc7spYKTO/0
# Z9YRRamhVWxxrIsBN5LrWh+18soXJ++EeSjzSYdgGWYPg16hL/7Aydx4Kz/WBTUm
# bGiiVUcE/I0aQU2U/0NzUiIFIW80SvxeDWn6I+hyVg/sdFSALP5JT7wAe8zTvsrI
# 2hMpEVLdStFAMqanFYqtwZU5FoAsoPZ7h1ElWmKLZkXk8ePuALztNY1yseO0Twdu
# eIGcIwItrlBYg1XpPz1+pMhGMVble6KHunaKo5K/ldOM0mQQT4Vjg6ZbzRIVRoDc
# ArQ5//0875jOUvJtYyc7Hl04jcmvjEIXC3HjkUYvgHEWL0QF/4f7vLAchaEZ839/
# 3GYOdqH5VVnZrUIBQB6DTaUILDCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkA
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
# T3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkRDMDAtMDVFMC1E
# OTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEw
# BwYFKw4DAhoDFQCJptLCZsE06NtmHQzB5F1TroFSBqCBgzCBgKR+MHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6Mg3SDAiGA8y
# MDIzMTAwNDE5MTEwNFoYDzIwMjMxMDA1MTkxMTA0WjB3MD0GCisGAQQBhFkKBAEx
# LzAtMAoCBQDoyDdIAgEAMAoCAQACAhFdAgH/MAcCAQACAhNKMAoCBQDoyYjIAgEA
# MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAI
# AgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAExNmbgvyJJqrFMVwdzx1uPxeJNW
# RFHp95MLdvwds78VB9xeDK+IBM7XZJ6lAaZjNPrWuL5xaUVAD6NQCScoG/aXNdvM
# PXKUYrEPUGu2+998UYlm/YwO4jWGX5FJnjQPu5Bt//IhkoE7hFyhbY+x2whF+u+7
# LHXrtksj5WagMLXVGF7t8Vxn04olW30PDy0cpoDEwo4ZZIQKA/ePH9dkE8/Q9zmL
# QjR+MAnmqrCs6q4kRiDBJTppawlJdDmn7S/KBW2qaf+u3lQU+Q+TGM98CBkT9WS8
# rrsTa1/AK/sSzATKfsBcsW3oeVtVJ9nFfAoUjbDE4XOHFrKoEbFeNcGx8p0xggQN
# MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdIh
# JDFKWL8tEQABAAAB0jANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
# CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCD+YCQ3ZLSCFuMg6cNKFpN4jkaa
# 6xBNiu2dorBD6i9AMDCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIMeAIJPf
# 30i9ZbOExU557GwWNaLH0Z5s65JFga2DeaROMIGYMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAHSISQxSli/LREAAQAAAdIwIgQgbQDYO6s7
# kxvwBMYBx9pSSDJYQrvNuJbO5RjIKWNnL4swDQYJKoZIhvcNAQELBQAEggIArIYt
# nAXbEPrQNl8P6VTJHfUZxyBRVi3GsT2lXVF7EXUa4rRuw26CVxws7FIUgr5PLvrX
# kae6V7OYc4RfC6nW4byZVDWzkwnKqeIjMF4SD4op1NNMBV5Er3jvcHrVxXKS/Bhx
# 1imf/6TcL9hebvSF4xOjps/HjF0Qt6C4JgP8fmYTwNepH+z9pBMpzvWBRLBDeoRD
# baupw0t5TUJQSldn3tKQO8VVdaJnfDs1AH3PBzBf4dasVBFittV/P4ZZjFqVFHOH
# Uv72qwGeELAgriOHBNP/Jm93TmPHv4CJEap0Wr9b69wRPzej0NrsAHxXAR79WuPo
# MCvCgNQ0xKGRfRfd02Nylcl1GzhpZ2PgqDDV7vJDhlr2Eo8tqyh3TQ+BW7AsRuH5
# Jq33LOj3Hc1wTO1lhOpRdY9AMEIcZ5pB7aQkr1motUUyyiMZnRZHU+E7f1kOVft0
# vZ29S7SDlYIiKrX2YpTdnp6KizP4shWkY0GFojyfNP270NVBQjKuriHG1miOQzkG
# KDC19vQdxvdUyVRuoDxs2/g1BQPHtEEVMmCMOM72Mk9Ri2Pmuu1/h2cv/xlNqYHj
# KAGJiMaLXHSFw+IPht/5e4QvmBZooVFADnig1P4hIfcczI+mRjC5y06cIfyRX2/6
# B51HwoevEkwuavHC4Pj7ig6a2PxEFHOQajhdS5E=
# SIG # End signature block
