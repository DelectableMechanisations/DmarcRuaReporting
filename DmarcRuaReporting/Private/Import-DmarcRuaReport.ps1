Function Import-DmarcRuaReport {
    [CmdletBinding()]
    Param (
        #Default:    .\Databases\UnprocessedXmlRuaReport
        [Parameter(Mandatory)]
        [System.String]
        $ImportDirectoryPath,

        #Default:    .\Databases\DMARCReportDatabase
        [Parameter(Mandatory)]
        [System.String]
        $ExportDirectoryPath,

        #Default:    .\Databases\ProcessedXmlRuaReport
        [Parameter(Mandatory)]
        [System.String]
        $CompletedDirectoryPath,

        #Default:    .\zz-ImportError
        [Parameter(Mandatory)]
        [System.String]
        $ErrorsProcessingDirectoryPath,

        #Default:    .\Databases\ProcessedXmlRuaReport-DB.xml
        [Parameter(Mandatory)]
        [System.String]
        $ProcessedXmlRuaReportDBFilePath,

        [Parameter(Mandatory)]
        [System.String]
        $IPToServerNameDatabaseFilePath,

        [System.String[]]
        $AcceptedFileExtension,

        #The maximum number of reports that can be kept in memory before they are paged to disk.
        [System.Int32]
        $MaxPageSize = 1000
    )

    Begin {
        Write-Debug -Message "Start processing RUA Report files from '$ImportDirectoryPath'."

        $writeProgressStopWatch  = [System.Diagnostics.Stopwatch]::StartNew()
        $writeProgressCounter    = 1

        #Create a temporary directory to store any files that generate an error when parsing.
        $tempErrorsProcessingDirectoryRoot = "$ErrorsProcessingDirectoryPath\ImportRuaXml_$(Get-Date -Format 'yyyy-MM-dd HHmmss')"
        Write-Debug -Message "Creating a temporary errors directory root '$tempErrorsProcessingDirectoryRoot'."
        New-Item -Path $tempErrorsProcessingDirectoryRoot -ItemType Directory -Force | Out-Null

        $dmarcMinDate = (Get-Date '1969-12-31 23:00:00 +00:00')


        if ((Get-Item -LiteralPath $IPToServerNameDatabaseFilePath).Length -eq 0) {
            Write-Debug -Message "  Total Keys: 0 (recreating)"
            $ipToServerNameDatabase = @{}

        } else {
            $ipToServerNameDatabase = Import-CliXml -LiteralPath $IPToServerNameDatabaseFilePath -ErrorAction Stop
            Write-Debug -Message "  Total Keys: $($IPToServerNameDatabase.Keys.Count)"
        }

        #Create a temporary database for all IPs that don't resolve to anything.
        $ipToNXDomainDatabase = @{}

        #Place a lock on the file to prevent other processes from accessing it.
        $fileLock = [System.IO.File]::Open($IPToServerNameDatabaseFilePath, 'Open', 'ReadWrite', 'None')

        $importFiles = @(Get-ChildItem -LiteralPath $ImportDirectoryPath -File -Recurse)

        $dmarcRuaReports = New-Object -TypeName System.Collections.ArrayList
    }

    Process {
        foreach ($file in $importFiles) {
            Write-Debug   -Message "  File: $($file.FullName)"

            #Only update the progress bar ever 100 Milliseconds, otherwise run time is > 100 slower
            if ($writeProgressStopWatch.Elapsed.TotalMilliseconds -ge 100) {
                $writeProgressParameters = @{
                    Activity        = 'Parsing XML report files...'
                    Status          = "File $writeProgressCounter of $($importFiles.Count)"
                    PercentComplete = ($writeProgressCounter/$importFiles.Count*100)
                }

                Write-Progress @writeProgressParameters
                $writeProgressStopWatch.Reset()
                $writeProgressStopWatch.Start()
            }

            #Confirm the file has an accepted file extension.
            if (Test-StringAgainstFilterList -String $file.Name -FilterList $AcceptedFileExtension) {

                <#
                    Create a temporary list that will store all processed records in a file.
                    If an error is encountered part way through processing the file, then we avoid polluting the main list with half processed items.
                #>
                $tempDmarcRuaReports = New-Object -TypeName System.Collections.ArrayList

                try {
                    [System.Xml.XmlDocument]$ruaReport = Get-Content -LiteralPath $file.FullName
                    $records = @($ruaReport.feedback.record)

                    $dateBegin = $dmarcMinDate.AddSeconds($ruaReport.feedback.report_metadata.date_range.begin)
                    $dateEnd   = $dmarcMinDate.AddSeconds($ruaReport.feedback.report_metadata.date_range.end)

                    foreach ($record in $records) {

                        #DNS Reverse Lookup (PTR) of Source IP
                        #-------------------------------------
                        $sourceIp = $record.row.source_ip
                        if ([String]::IsNullOrWhiteSpace($sourceIp)) {
                            $serverName = '<Not recorded>'

                        #Check the database to see if the IP to Server Name (DNS PTR) translation already exists.
                        } elseif ($ipToServerNameDatabase.ContainsKey($sourceIp)) {
                            $serverName = $ipToServerNameDatabase.Item($sourceIp)

                        #Check the temporary NXDOMAIN database to see if the IP to Server Name (DNS PTR) translation has already been tried but has failed.
                        } elseif ($ipToNXDomainDatabase.ContainsKey($sourceIp)) {
                            $serverName = $ipToNXDomainDatabase.Item($sourceIp)

                        #If the IP has never been encountered before then try and resolve it.
                        } else {
                            $resolutionTest = Resolve-IpToPtr -IPAddress $sourceIp
                            switch ($resolutionTest.ResolvedToPtr) {
                                $true  {$ipToServerNameDatabase[$sourceIp] = $resolutionTest.ResolvedPtr}
                                $false {$ipToNXDomainDatabase[$sourceIp]   = $resolutionTest.ResolvedPtr}
                            }
                            $serverName = $resolutionTest.ResolvedPtr
                        }

                        $dmarcAlignmentResult = Test-DmarcAlignment -HeaderFrom $record.identifiers.header_from -AuthResult $record.auth_results

                        [VOID]$tempDmarcRuaReports.Add(([PSCustomObject]@{
                            'SourceReport_FileName'              = $file.FullName
                            'DestinationReport_ParentFolderName' = $dateBegin.ToString('yyyy-MM-dd')
                            'DestinationReport_FileName'         = "DMARC_Processed_Data_$($record.identifiers.header_from).csv"
                            'Metadata_OrgName'                   = $ruaReport.feedback.report_metadata.org_name
                            'Metadata_Email'                     = $ruaReport.feedback.report_metadata.email
                            'Metadata_ReportID'                  = $ruaReport.feedback.report_metadata.report_id
                            'Metadata_DateBegin'                 = $dateBegin.ToString('yyyy-MM-dd HH:mm:ss')
                            'Metadata_DateEnd'                   = $dateEnd.ToString('yyyy-MM-dd HH:mm:ss')
                            'Published_Policy_Domain'            = $ruaReport.feedback.policy_published.domain
                            'Published_Policy_adkim'             = $ruaReport.feedback.policy_published.adkim
                            'Published_Policy_adspf'             = $ruaReport.feedback.policy_published.aspf
                            'Published_Policy_Policy'            = $ruaReport.feedback.policy_published.p
                            'Published_Policy_SubdomainPolicy'   = $ruaReport.feedback.policy_published.sp
                            'Published_Policy_Percentage'        = $ruaReport.feedback.policy_published.pct
                            'Published_Policy_ReportFormat'      = $ruaReport.feedback.policy_published.fo
                            'Row_SourceIP'                       = $record.row.source_ip
                            'Row_Count'                          = $record.row.count
                            'Row_PolicyEvaluated_Disposition'    = $record.row.policy_evaluated.disposition
                            'Row_PolicyEvaluated_DKIM'           = $record.row.policy_evaluated.dkim
                            'Row_PolicyEvaluated_SPF'            = $record.row.policy_evaluated.spf
                            'Identifiers_Envelope_To'            = $record.identifiers.envelope_to
                            'Identifiers_Envelope_From'          = $record.identifiers.envelope_from
                            'Identifiers_Header_From'            = $record.identifiers.header_from
                            'AuthResults_DKIM_Domain'            = [String]($record.auth_results.dkim.domain)
                            'AuthResults_DKIM_Selector'          = [String]($record.auth_results.dkim.selector)
                            'AuthResults_DKIM_Result'            = [String]($record.auth_results.dkim.result)
                            'AuthResults_DKIM_HumanResult'       = [String]($record.auth_results.dkim.human_result)
                            'AuthResults_SPF_Domain'             = $record.auth_results.spf.domain
                            'AuthResults_SPF_Scope'              = $record.auth_results.spf.scope
                            'AuthResults_SPF_Result'             = $record.auth_results.spf.result
                            'Date Begin'                         = $dateBegin.ToString('yyyy-MM-dd HH:mm:ss')
                            'Date End'                           = $dateEnd.ToString('yyyy-MM-dd HH:mm:ss')
                            'From Domain'                        = $record.identifiers.header_from
                            'DMARC Alignment'                    = $dmarcAlignmentResult.DmarcAlignment
                            'IP'                                 = $record.row.source_ip
                            'Server (from PTR)'                  = $serverName
                            'Email Count'                        = $record.row.count
                            'Policy Applied'                     = $record.row.policy_evaluated.disposition
                            'SPF (DMARC)'                        = $dmarcAlignmentResult.SpfAuthResult
                            'SPF (Raw result)'                   = $dmarcAlignmentResult.SpfResult
                            'SPF (Mail From)'                    = $dmarcAlignmentResult.SpfDomain
                            'DKIM (DMARC)'                       = $dmarcAlignmentResult.DkimAuthResult
                            'DKIM (Raw result)'                  = $dmarcAlignmentResult.DkimResult
                            'DKIM (domain)'                      = $dmarcAlignmentResult.DkimDomain
                            'DKIM (selector)'                    = $dmarcAlignmentResult.DkimSelector
                            'Reported By'                        = $ruaReport.feedback.report_metadata.org_name
                        }))
                    }

                    #If no issues encountered when processing reports in this file, add it to the overall list.
                    if ($tempDmarcRuaReports.Count -gt 1) {
                        [VOID]$dmarcRuaReports.AddRange($tempDmarcRuaReports)

                    } elseif ($tempDmarcRuaReports.Count -eq 1) {
                        [VOID]$dmarcRuaReports.Add($tempDmarcRuaReports[0])
                    }


                    #Archive the original file once completed processing.
                    Write-Debug -Message "    Successfully processed all records in '$($file.FullName)'."
                    Write-Debug -Message "    Moving original file into '$CompletedDirectoryPath'."
                    Move-Item -LiteralPath $file.FullName -Destination $CompletedDirectoryPath

                    #Export reports to disk once the threshold has been reached.
                    if ($dmarcRuaReports.Count -ge $MaxPageSize) {
                        Export-DmarcRuaReport `
                        -DmarcRuaReport      $dmarcRuaReports `
                        -ExportDirectoryPath $ExportDirectoryPath

                        $dmarcRuaReports.Clear()
                    }

                } catch {
                    #If issues are encountered whilst processing the file, remove it from the list of processed
                    Remove-FileFromHashDatabase `
                    -FileHashDatabasePath $ProcessedXmlRuaReportDBFilePath `
                    -FilePath             $file.FullName


                    Write-Error -Message @"
$($file.Name)
    Failed to parse into XML with error message: $_
    Moving into '$tempErrorsProcessingDirectoryRoot'.
"@
                    Move-Item -LiteralPath $file.FullName -Destination $tempErrorsProcessingDirectoryRoot
                }

            #If the file does not have an accepted file extension then leave in the present location.
            } else {
                Write-Debug -Message "    File has an invalid extension."
                Write-Debug -Message "    Moving original file into '$tempErrorsProcessingDirectoryRoot'."
                Move-Item -LiteralPath $file.FullName -Destination $tempErrorsProcessingDirectoryRoot
            }
            $writeProgressCounter++
        }

        #Export any reports still in memory.
        if ($dmarcRuaReports.Count -ge 1) {
            Export-DmarcRuaReport `
            -DmarcRuaReport      $dmarcRuaReports `
            -ExportDirectoryPath $ExportDirectoryPath
        }
    }

    End {
        #Remove the file lock and export the data to the database.
        $fileLock.Close()
        $ipToServerNameDatabase | Export-CliXml -LiteralPath $IPToServerNameDatabaseFilePath -ErrorAction Stop

        #Remove the temporary error processing directory if it is empty.
        Remove-EmptyDirectory -LiteralPath $tempErrorsProcessingDirectoryRoot -IncludeRoot
    }
}
