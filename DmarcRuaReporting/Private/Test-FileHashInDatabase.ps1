Function Test-FileHashInDatabase {
    [OutputType([System.Boolean])]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [System.Collections.Hashtable]
        $FileHashDatabase,

        [Parameter(Mandatory)]
        [System.String]
        $FileHash
    )

    $searchResult = $FileHashDatabase.ContainsKey($FileHash)
    Write-Output -InputObject $searchResult
}
