Function Test-DmarcAlignment {
    Param (
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [System.String]
        $HeaderFrom,

        [Parameter(Mandatory)]
        [System.Object]
        $AuthResult
    )

    #If the HeaderFrom parameter is empty then no alignment check can be performed (there's nothing to align on).
    if ([System.String]::IsNullOrWhiteSpace($HeaderFrom)) {
        $overallResult = [PSCustomObject]@{
            DkimAuthResult  = 'UNKNOWN'
            DkimDomain      = 'UNKNOWN'
            DkimSelector    = 'UNKNOWN'
            DkimResult      = 'UNKNOWN'
            DkimHumanResult = 'UNKNOWN'
            SpfAuthResult   = 'UNKNOWN'
            SpfDomain       = 'UNKNOWN'
            SpfScope        = 'UNKNOWN'
            SpfResult       = 'UNKNOWN'
            DmarcAlignment  = "UNKNOWN - Record report is missing mandatory 'header_from' field"
        }


    #Evaluate alignment if the HeaderFrom parameter is present.
    } else {
        $dkimAuthRawResult = Test-DmarcDkimAuthResult -HeaderFrom $HeaderFrom -AuthResult $AuthResult
        $spfAuthRawResult  = Test-DmarcSpfAuthResult  -HeaderFrom $HeaderFrom -AuthResult $AuthResult

        $overallResult = [PSCustomObject]@{
            DkimAuthResult  = $dkimAuthRawResult.DmarcDkimAlignment
            DkimDomain      = $dkimAuthRawResult.DkimDomain
            DkimSelector    = $dkimAuthRawResult.DkimSelector
            DkimResult      = $dkimAuthRawResult.DkimResult
            DkimHumanResult = $dkimAuthRawResult.DkimHumanResult
            SpfAuthResult   = $spfAuthRawResult.DmarcSpfAlignment
            SpfDomain       = $spfAuthRawResult.SpfDomain
            SpfScope        = $spfAuthRawResult.SpfScope
            SpfResult       = $spfAuthRawResult.SpfResult
            DmarcAlignment  = ''
        }

        switch ($overallResult) {
            {$_.DkimAuthResult -like 'DKIM Aligned*' -and $_.SpfAuthResult -like 'SPF Aligned*'}    {$overallResult.DmarcAlignment = 'Aligned (DKIM and SPF)'; break}
            {$_.DkimAuthResult -like 'DKIM Aligned*'}                                               {$overallResult.DmarcAlignment = 'Aligned (DKIM only)'; break}
            {$_.SpfAuthResult  -like 'SPF Aligned*'}                                                {$overallResult.DmarcAlignment = 'Aligned (SPF only)'; break}
            Default                                                                                 {$overallResult.DmarcAlignment = 'Fail (unauthenticated)'}
        }
    }

    Write-Output -InputObject $overallResult
}
