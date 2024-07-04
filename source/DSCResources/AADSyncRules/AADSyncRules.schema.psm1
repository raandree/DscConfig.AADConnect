configuration AADSyncRules {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable[]]
        $Items
    )

    Import-DscResource -ModuleName AADConnectDsc
    Import-DscResource -ModuleName PSDesiredStateConfiguration

    foreach ($item in $Items)
    {
        if (-not $item.ContainsKey('Ensure'))
        {
            $item.Ensure = 'Present'
        }

        $executionName = ($item.ConnectorName + '__' + $item.Name) -replace '[\s(){}/\\:-]', '_'
        (Get-DscSplattedResource -ResourceName AADSyncRule -ExecutionName $executionName -Properties $item -NoInvoke).Invoke($item)
    }
}
