configuration AADConnectDirectoryExtensionAttributes {
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

        $executionName = ($item.Name + '__' + $item.AssignedObjectClass) -replace '[\s(){}/\\:-]', '_'
        (Get-DscSplattedResource -ResourceName AADConnectDirectoryExtensionAttribute -ExecutionName $executionName -Properties $item -NoInvoke).Invoke($item)
    }
}
