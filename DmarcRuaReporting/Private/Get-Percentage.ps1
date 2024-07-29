Function Get-Percentage {
    Param (
        [System.Double]
        $Count,

        [System.Double]
        $Total,

        [System.Int32]
        $DecimalPlaces = 2
    )
    $percentage = ($Count/$Total)*100

    $roundedPercentage = Get-RoundNumber -Number $percentage -DecimalPlaces $DecimalPlaces
    Write-Output -InputObject $roundedPercentage
}
