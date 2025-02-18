<#
    .SYNOPSIS
        Generates a CSV based report using information stored in the DMARC Report Database.

    .DESCRIPTION
        The Start-DmarcRuaReport function is used to create human readable reports from the data stored in the DMARC database.
        These reports use the CSV format and are intended to be reviewed within a spreadsheet viewer like Microsoft Excel.

    .PARAMETER Path
        Specifies the directory root path to the DMARC RUA Report Database.
        Defaults to . (the current location).

    .PARAMETER ReportType
        Specifies the type of report to generate.
        The acceptable values for this parameter are:
            - All: Generates a single CSV report file containing all successful and failed DMARC alignments for all domains.
            - AllFailedAlignment: Generates a single CSV report file containing all failed DMARC alignments for all domains.
            - PerDomain: Generates one CSV report file per domain containing all successful and failed DMARC alignments.
            - AllSummary: Generates a single CSV report file that gives a high level summary of DMARC aligned emails vs DMARC not aligned emails on a per domain basis.
        
        The default vaule for this paramter is 'PerDomain' and multiple values can be specified.
        
    .PARAMETER ReportStyle
        Specifies the style of report to generate and refers to the number of columns included in the report.
        The acceptable values for this parameter are:
            - Raw: All columns in the DMARC database are included in the report output. These columns contain data taken directly from the original XML reports and are used to construct the more human friendly columns included in the Compact report.
            - Compact: Only includes the human friendly columns in the report output.

        The default vaule for this paramter is 'Compact'.
    
    .PARAMETER StartDate
        Specifies the start date and time of the date range.
        It can take one of these forms:
            - A System.DateTime object (e.g. (Get-Date).AddMonths(-5))
            - A System.String object (e.g. '2024-05-10 20:00:00')
        
        The default vaule for this paramter is 3 months ago.
         
    .PARAMETER DomainFilter
        Specifies a file filter to apply when determining which domains to include in the report.
        When using this parameter, ensure it starts and ends with the '*' wildcard character.
        i.e. '*primarydomain*' or '*domain*'

        The default value for this paramter is '*' (i.e. all files).

    .EXAMPLE
        Start-DmarcRuaReport

        Generates DMARC RUA reports using the default parameter values.

    .EXAMPLE
        Start-DmarcRuaReport `
        -ReportType  PerDomain, AllSummary `
        -StartDate   (Get-Date).AddMonths(-6)

        Generates 2 sets of DMARC RUA reports: 1 using the 'PerDomain' type and 1 using the 'AllSummary' type.
        These reports include all data collected over the last 6 months.

    .EXAMPLE
        Start-DmarcRuaReport `
        -ReportType   AllSummary `
        -StartDate    '2024-05-01' `
        -DomainFilter *primary*

        Generates an 'AllSummary' report including all data collected since 2024-05-01 that were for domains that contain the word 'primary'.
#>
Function Start-DmarcRuaReport {
    [CmdletBinding()]
    Param (
        [ValidateScript({Test-Path -LiteralPath $_ -PathType Container})]
        [System.String]
        $Path = (Get-Item -Path .).FullName,

        [ValidateSet('All', 'AllFailedAlignment', 'PerDomain', 'AllSummary')]
        [System.String[]]
        $ReportType = 'PerDomain',

        [ValidateSet('Raw', 'Compact')]
        [System.String]
        $ReportStyle = 'Compact',

        [System.DateTime]
        $StartDate = (Get-Date).AddMonths(-3),

        [System.String]
        $DomainFilter = '*'
    )

    $functionStartDate  = Get-Date
    Write-Verbose -Message "Report started: $($functionStartDate.ToString('yyyy-MM-dd HH:mm:ss'))"

    $testDmarcRuaReportDatabase = Test-DmarcRuaReportDatabase -Path $Path
    if ($testDmarcRuaReportDatabase -eq $false) {
        throw "Could not find a valid DMARC RUA Report Database in the path '$Path'."
    }

    #Import all paths in the DevicePath format in order to get around the 256 character path limit in Windows Explorer.
    $drrConfigData = Get-DmarcRuaReportDatabase -PathFormat 'DevicePath'

    #Search the Output Directory Path for existing reports and archive any that are found.
    $existingReports = @(Get-ChildItem -LiteralPath $drrConfigData.ReportOutputDirectory -Filter 'DMARC Report - *.csv')
    if ($existingReports.Count -gt 0) {
        Write-Verbose -Message "Removing old reports from '$($drrConfigData.ReportOutputDirectory)'."
        $existingReports | Remove-Item -Force
    }


    $summaryData = New-Object -TypeName System.Collections.ArrayList

    #Select the column filters based on the Report Style.
    switch ($ReportStyle) {
        #Selects the most useful columns in the report, ignoring the columns with a '_' in them that are from the raw XML files.
        'Compact' {$selectObjectParams = @{Property = '*'; ExcludeProperty = '*_*'}; break}

        #Selects all columns in the report.
        'Raw'     {$selectObjectParams = @{Property = '*'};                          break}
    }

    $dmarcReportDateDirectories = @(Get-ChildItem -LiteralPath $drrConfigData.DmarcReportDatabaseFiles)
    $dmarcReportDatabaseFiles = New-Object -TypeName System.Collections.ArrayList
    foreach ($directory in $dmarcReportDateDirectories) {
        Write-Debug -Message "Processing: $($directory.FullName)"

        #Include only the reports that are later than $StartDate
        $directoryDate = Get-Date $directory.Name
        if ($directoryDate -gt $StartDate) {
            Write-Debug -Message "  Including directory because it is newer than the value specified in -StartDate ($($StartDate.ToString('yyyy-MM-dd HH:mm:ss')))."

            #Include only the reports that match the $DomainFilter.
            $filesMatchingDomainFilter = @(Get-ChildItem -LiteralPath $directory.FullName -Filter $DomainFilter)
            foreach ($file in $filesMatchingDomainFilter) {
                Write-Debug -Message "  Report file matches -DomainFilter '$DomainFilter': $($file.Name)"
                [VOID]$dmarcReportDatabaseFiles.Add($file.FullName)
            }

        } else {
            Write-Debug -Message "Excluding '$($directory.FullName)' because it is older than the value specified in -StartDate ($($StartDate.ToString('yyyy-MM-dd HH:mm:ss')))."
        }
    }

    $writeProgressStopWatch  = [System.Diagnostics.Stopwatch]::StartNew()
    $writeProgressCounter    = 1
    foreach ($dmarcReportDatabaseFile in $dmarcReportDatabaseFiles) {
        #Only update the progress bar ever 100 Milliseconds, otherwise run time is > 100 slower
        if ($writeProgressStopWatch.Elapsed.TotalMilliseconds -ge 100) {
            $writeProgressParameters = @{
                Activity        = 'Processing Database Files...'
                Status          = "File $writeProgressCounter of $($dmarcReportDatabaseFiles.Count)"
                PercentComplete = ($writeProgressCounter/$dmarcReportDatabaseFiles.Count*100)
            }

            Write-Progress @writeProgressParameters
            $writeProgressStopWatch.Reset()
            $writeProgressStopWatch.Start()
        }

        $dmarcReportData = @(Import-Csv -LiteralPath $dmarcReportDatabaseFile | Sort-Object -Property 'Date Begin')

        #Add a Sender Category to the email using the mappings in the Get-SenderCategory function.
        foreach ($item in $dmarcReportData) {
            $item | Add-Member -MemberType NoteProperty -Name 'Sender Category' -Value ''

            $senderCategory = Get-SenderCategory -SenderString @(
                $item.'Server (from PTR)',
                $item.'SPF (Mail From)'
            )

            if ($null -ne $senderCategory) {
                $item.'Sender Category' = $senderCategory
            }
        }

        switch ($ReportType) {
            #Complete export of the DMARC database into a single .csv file.
            'All' {
                $exportPath = "$($drrConfigData.ReportOutputDirectory)\DMARC Report - All.csv"
                $dmarcReportData | Select-Object @selectObjectParams | Export-Csv -LiteralPath $exportPath -Append -NoTypeInformation
            }

            #Complete export of all reported failed alignments in the DMARC database into a single .csv file.
            'AllFailedAlignment' {
                $exportPath = "$($drrConfigData.ReportOutputDirectory)\DMARC Report - All Failed Alignment.csv"
                $dmarcReportData | Select-Object @selectObjectParams | Where-Object {$_.'DMARC Alignment' -like 'Fail*'} |
                Export-Csv -LiteralPath $exportPath -Append -NoTypeInformation
            }

            #Generate 1 report file per domain.
            'PerDomain' {
                $domainGroups = @($dmarcReportData | Select-Object @selectObjectParams | Group-Object -Property 'From Domain')
                foreach ($domainGroup in $domainGroups) {
                    $exportPath = "$($drrConfigData.ReportOutputDirectory)\DMARC Report - Per Domain - $($domainGroup.Name).csv"
                    $domainGroup.Group | Export-Csv -LiteralPath $exportPath -Append -NoTypeInformation
                }
            }

            'AllSummary' {
                $dmarcReportData | Select-Object -Property @(
                    @{Label = 'DateMonth';       Expression = {$_.'Date Begin'.ToString('yyyy-MM')}},
                    @{Label = 'DateYear';        Expression = {$_.'Date Begin'.ToString('yyyy')}},
                    @{Label = 'Domain';          Expression = {$_.'From Domain'}},
                    @{Label = 'DMARC Alignment'; Expression = {Test-DmarcAlignmentSummary -DmarcAlignment $_.'DMARC Alignment'}}
                    @{Label = 'Total Emails';    Expression = {$_.'Email Count'}}
                ) | ForEach-Object {[VOID]$summaryData.Add($_)}
            }
        }
        $writeProgressCounter++
    }

    if ($ReportType -eq 'AllSummary') {
        $dmarcReportDomains = @($summaryData | Group-Object -Property 'Domain')
        $domainSummaries = New-Object -TypeName System.Collections.ArrayList
        foreach ($dmarcReportDomain in $dmarcReportDomains) {
            $domainGroupedByAlignment = @($dmarcReportDomain.Group | Group-Object -Property 'DMARC Alignment')

            #All data.
            $aligned     = $domainGroupedByAlignment | Where-Object {$_.Name -eq 'Aligned'}    | Select-Object -ExpandProperty Group | Measure-Object -Property 'Total Emails' -Sum
            $notAligned  = $domainGroupedByAlignment | Where-Object {$_.Name -eq 'NotAligned'} | Select-Object -ExpandProperty Group | Measure-Object -Property 'Total Emails' -Sum

            [VOID]$domainSummaries.Add(([PSCustomObject]@{
                Domain                     = $dmarcReportDomain.Name
                TotalAlignedEmails         = [Int32]($aligned.Sum)
                TotalNotAlignedEmails      = [Int32]($notAligned.Sum)
                TotalAlignedPercentage     = (Get-Percentage -Count $aligned.Sum -Total ($aligned.Sum + $notAligned.Sum))
            }))
        }
        $exportPath = "$($drrConfigData.ReportOutputDirectory)\DMARC Report - Summary.csv"
        $domainSummaries | Sort-Object -Property Domain | Export-Csv -LiteralPath $exportPath -NoTypeInformation

    }

    Write-Verbose -Message "Report export completed in $(New-TimeSpan -Start $functionStartDate)"
}
