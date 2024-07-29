Function Get-RoundNumber {
    Param (
        [System.Double]
        $Number,

        [System.Int32]
        $DecimalPlaces = 2
    )

    $roundNumber = [System.Math]::Round($Number, $DecimalPlaces)
    Write-Output -InputObject $roundNumber
}
