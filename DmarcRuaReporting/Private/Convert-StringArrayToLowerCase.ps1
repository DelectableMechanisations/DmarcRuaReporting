<#
    .SYNOPSIS
        Takes an array of strings and converts each element in the array to
        all lowercase characters.

    .PARAMETER Array
        The array of System.String objects to convert into lowercase strings.

    .NOTES
        Sourced from PowerShell module ExchangeDsc version 2.0.0
#>
function Convert-StringArrayToLowerCase
{
    [CmdletBinding()]
    [OutputType([System.String[]])]
    param
    (
        [Parameter()]
        [System.String[]]
        $Array
    )

    [System.String[]] $arrayOut = New-Object -TypeName 'System.String[]' -ArgumentList $Array.Count

    for ($i = 0; $i -lt $Array.Count; $i++)
    {
        $arrayOut[$i] = $Array[$i].ToLower()
    }

    return $arrayOut
}
