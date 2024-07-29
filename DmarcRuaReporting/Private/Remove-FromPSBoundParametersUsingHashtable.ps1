<#
    .SYNOPSIS
        Takes $PSBoundParameters from another function, and modifies it based
        on the contents of the ParamsToRemove or ParamsToKeep parameters. If
        ParamsToRemove is specified, it will remove each param. If ParamsToKeep
        is specified, everything but those params will be removed. If both
        ParamsToRemove and ParamsToKeep are specified, the function will throw
        an exception.

    .PARAMETER PSBoundParametersIn
        The $PSBoundParameters Hashtable from the calling function.

    .PARAMETER ParamsToKeep
        A String array containing the list of parameter names to keep in the
        given PSBoundParametersIn HashTable.

    .PARAMETER ParamsToRemove
        A String array containing the list of parameter names to remove in the
        given PSBoundParametersIn HashTable.

    .NOTES
        Sourced from PowerShell module ExchangeDsc version 2.0.0
#>
function Remove-FromPSBoundParametersUsingHashtable
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        $PSBoundParametersIn,

        [Parameter()]
        [System.String[]]
        $ParamsToKeep,

        [Parameter()]
        [System.String[]]
        $ParamsToRemove
    )

    if ($ParamsToKeep.Count -gt 0 -and $ParamsToRemove.Count -gt 0)
    {
        throw 'Remove-FromPSBoundParametersUsingHashtable does not support using both ParamsToKeep and ParamsToRemove'
    }

    if ($ParamsToKeep.Count -gt 0)
    {
        $ParamsToKeep = $ParamsToKeep.ToLower()

        $lowerParamsToKeep = Convert-StringArrayToLowerCase -Array $ParamsToKeep

        foreach ($key in $PSBoundParametersIn.Keys)
        {
            if (!($lowerParamsToKeep.Contains($key.ToLower())))
            {
                $ParamsToRemove += $key
            }
        }
    }

    if ($ParamsToRemove.Count -gt 0)
    {
        foreach ($param in $ParamsToRemove)
        {
            $PSBoundParametersIn.Remove($param) | Out-Null
        }
    }
}
