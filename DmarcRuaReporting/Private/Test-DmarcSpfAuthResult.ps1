Function Test-DmarcSpfAuthResult {
    Param (
        [Parameter(Mandatory)]
        [System.String]
        $HeaderFrom,

        [System.Object]
        $AuthResult
    )

    $overallResult = [PSCustomObject]@{
        DmarcSpfAlignment = ''
        SpfDomain         = $AuthResult.spf.domain
        SpfScope          = $AuthResult.spf.scope
        SpfResult         = $AuthResult.spf.result
    }
    switch ($AuthResult.spf) {
        {$_.result -eq 'pass' -and $_.domain -eq $HeaderFrom}       {$overallResult.DmarcSpfAlignment = 'SPF Aligned (Strict)'; break}
        {$_.result -eq 'pass' -and $_.domain -like "*.$HeaderFrom"} {$overallResult.DmarcSpfAlignment = 'SPF Aligned (Relaxed)'; break}
        {$_.result -eq 'pass' -and $_.domain -ne $HeaderFrom}       {$overallResult.DmarcSpfAlignment = 'SPF Not Aligned (check pass)'; break}
        Default                                                     {$overallResult.DmarcSpfAlignment = 'SPF Not Aligned (check fail)'; break}
    }

    Write-Output -InputObject $overallResult
}
