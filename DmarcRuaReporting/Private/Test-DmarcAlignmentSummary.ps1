Function Test-DmarcAlignmentSummary {
    Param (
        [Parameter(Mandatory)]
        [System.String]
        $DmarcAlignment
    )

    if ($DmarcAlignment -like 'Aligned*') {
        Write-Output -InputObject 'Aligned'

    } else {
        Write-Output -InputObject 'NotAligned'
    }
}
