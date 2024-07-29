Function Test-DmarcDkimAuthResult {
    Param (
        [Parameter(Mandatory)]
        [System.String]
        $HeaderFrom,

        [System.Object]
        $AuthResult
    )

    $dmarcDkimAlignments = New-Object -TypeName System.Collections.ArrayList
    foreach ($dkimAuthResult in $AuthResult.dkim) {
        $overallResult = [PSCustomObject]@{
            DmarcDkimAlignment = ''
            Level              = 0
            DkimDomain         = $dkimAuthResult.domain
            DkimSelector       = $dkimAuthResult.selector
            DkimResult         = $dkimAuthResult.result
            DkimHumanResult    = $dkimAuthResult.human_result
        }

        switch ($dkimAuthResult) {
            {$_.result -eq 'pass' -and $_.domain -eq $HeaderFrom}       {$overallResult.Level = 5; $overallResult.DmarcDkimAlignment = 'DKIM Aligned (Strict)';                      break}
            {$_.result -eq 'pass' -and $_.domain -like "*.$HeaderFrom"} {$overallResult.Level = 4; $overallResult.DmarcDkimAlignment = 'DKIM Aligned (Relaxed)';                     break}
            {$_.result -ne 'pass' -and $_.domain -eq $HeaderFrom}       {$overallResult.Level = 3; $overallResult.DmarcDkimAlignment = 'DKIM Not Aligned (Strict with check fail)';  break}
            {$_.result -ne 'pass' -and $_.domain -like "*.$HeaderFrom"} {$overallResult.Level = 2; $overallResult.DmarcDkimAlignment = 'DKIM Not Aligned (Relaxed with check fail)'; break}
            {$_.result -eq 'pass' -and $_.domain -ne $HeaderFrom}       {$overallResult.Level = 1; $overallResult.DmarcDkimAlignment = 'DKIM Not Aligned (check pass)';              break}
            Default                                                     {$overallResult.Level = 0; $overallResult.DmarcDkimAlignment = 'DKIM Not Aligned (check fail)'                    }
        }

        [VOID]$dmarcDkimAlignments.Add($overallResult)
    }

    #Select the DKIM auth result with the highest level of alignment and output this result.
    $bestDmarcDkimAlignmentResult = $dmarcDkimAlignments | Sort-Object -Property Level | Select-Object -Last 1
    $bestDmarcDkimAlignmentResult | Select-Object -Property * -ExcludeProperty Level | Write-Output
}
