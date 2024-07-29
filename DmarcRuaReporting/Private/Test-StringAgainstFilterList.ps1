Function Test-StringAgainstFilterList {
    Param (
        [System.String]
        $String,

        [System.String[]]
        $FilterList = @(
            '*.7z',
            '*.gz',
            '*.rar',
            '*.tar',
            '*.zip'
        )
    )

    $matchesFilter = $false
    foreach ($filter in $FilterList) {
        if ($String -like $filter) {
            $matchesFilter = $true
        }
    }

    Write-Output -InputObject $matchesFilter
}
