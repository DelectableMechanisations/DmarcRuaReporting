Function Get-SenderCategory {
    Param (
        [System.String[]]
        $SenderString
    )

    $senderCategory = $null
    foreach ($string in $SenderString) {
        switch ($string) {
            {[Regex]::IsMatch($_, '\.salesforce\.com$', 1)}                  {$senderCategory = 'SalesForce';        $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'cmail[0-9]{0,2}\.com$', 1)}               {$senderCategory = 'Campaign Monitor';  $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'outbound\.createsend\.com$', 1)}          {$senderCategory = 'Campaign Monitor';  $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'outbound\.protection\.outlook\.com$', 1)} {$senderCategory = 'Exchange Online';   $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'amazonses\.com$', 1)}                     {$senderCategory = 'Amazon SES';        $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'pphosted\.com$', 1)}                      {$senderCategory = 'ProofPoint';        $foundMatch = $true; break}
            {[Regex]::IsMatch($_, '\.google\.com$', 1)}                      {$senderCategory = 'Google Mail';       $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'xtra\.co\.nz$', 1)}                       {$senderCategory = 'Spark Xtra Mail';   $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'mcsv\.net$', 1)}                          {$senderCategory = 'MailChimp';         $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'mcdlv\.net$', 1)}                         {$senderCategory = 'MailChimp';         $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'rsgsv\.net$', 1)}                         {$senderCategory = 'MailChimp';         $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'mctxapp\.net$', 1)}                       {$senderCategory = 'MailChimp';         $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'aotal\.cloud$', 1)}                       {$senderCategory = 'AOTAL (HR System)'; $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'mtasv\.net$', 1)}                         {$senderCategory = 'Postmark';          $foundMatch = $true; break}
            {[Regex]::IsMatch($_, 'smtp2go\.com$', 1)}                       {$senderCategory = 'SMTP2Go';           $foundMatch = $true; break}
        }

        if ($foundMatch) {
            break
        }
    }

    Write-Output -InputObject $senderCategory
}
