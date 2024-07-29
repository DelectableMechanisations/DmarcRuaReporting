Function Start-DmarcRuaMailboxScrape {
    [CmdletBinding()]
    Param (
        [ValidateScript({ Test-Path -Path $_ -PathType Container })]
        [System.String]
        $Path = (Get-Item -Path .).FullName,

        [System.String[]]
        $DisplayName,

        [System.String]
        $SourceFolder = 'Inbox',

        [System.String]
        $DestinationFolder = '_DmarcProcessed',

        [System.Int32]
        $ProcessingLimit = 5000
    )

    $testDmarcRuaReportDatabase = Test-DmarcRuaReportDatabase -Path $Path
    if ($testDmarcRuaReportDatabase -eq $false) {
        throw "Could not find a valid DMARC RUA Report Database in the path '$Path'."
    }

    $drrConfigData = Get-DmarcRuaReportDatabase

    #If mailbox DisplayName is specified as as parameter, this takes precedence.
    if ($PSBoundParameters.ContainsKey('DisplayName')) {
        $mailboxScrapeSources = @(Get-DmarcRuaMailboxScrapeSource -DisplayName $DisplayName)

    #Otherwise take the mailbox scrape source from the value in the config.json file.
    } elseif ($drrConfigData.OutlookMailboxesToScrape.Count -gt 0) {
        $mailboxScrapeSources = @(Get-DmarcRuaMailboxScrapeSource -DisplayName $drrConfigData.OutlookMailboxesToScrape)

    #Stop gracefully if neither of the previous conditions evaluate to $true.
    } else {
        Write-Warning -Message "The -DisplayName parameter was not specified and there are no Outlook Mailboxes listed as scrape sources in 'config.json'."
        return
    }

    #Confirm all mailbox scrape sources are valid.
    foreach ($mailboxScrapeSource in $mailboxScrapeSources) {
        if ($mailboxScrapeSource.IsValidMailboxScrapeSource -eq $true) {
            Write-Verbose -Message "Validated '$($mailboxScrapeSource.DisplayName)' as a mailbox scrape source."

        } else {
            throw "Mailbox '$($mailboxScrapeSource.DisplayName)' is not a valid mailbox scrape source."
        }
    }

    $mailboxScrapeDirectoryRelative = Join-Path -Path $drrConfigData.ImportSourceDirectory -ChildPath "$(Get-Date -Format 'yyyy-MM-dd HHmmss')_MailboxScrape"
    $mailboxScrapeDirectory = (New-Item -Path $mailboxScrapeDirectoryRelative -ItemType Directory -Force).FullName

    $mailboxScrapeSourceDisplayNames = @($mailboxScrapeSources | Select-Object -ExpandProperty DisplayName)

    $outlook = New-Object -ComObject 'Outlook.Application' -ErrorAction 'Stop'
    $outlookStores = @($outlook.GetNamespace('MAPI').Stores | Where-Object {$mailboxScrapeSourceDisplayNames -contains $_.DisplayName})

    $writeProgressStopWatch  = [System.Diagnostics.Stopwatch]::StartNew()

    $attachmentCounter = 1
    foreach ($outlookStore in $outlookStores) {
        $writeProgressCounter = 1
        $outlookSourceFolder = $outlookStore.GetRootFolder().Folders | Where-Object {$_.Name -eq $SourceFolder}
        $outlookItems  = @($outlookSourceFolder.Items)

        #Find the destination folder. If doesn't exist then create it.
        $outlookDestinationFolder = $outlookStore.GetRootFolder().Folders | Where-Object {$_.Name -eq $DestinationFolder}
        if ($null -eq $outlookDestinationFolder) {
            $outlookDestinationFolder = $outlookStore.GetRootFolder().Folders.Add($DestinationFolder)
        }

        Write-Verbose -Message "Start processing $($outlookStore.DisplayName)\$SourceFolder"
        Write-Debug   -Message "  Total items to process: $($outlookItems.Count)"

        $outlookItemCounter = 0
        Do {
            $outlookItem = $outlookItems[$outlookItemCounter]

            Write-Verbose -Message "    [$outlookItemCounter] Received $($outlookItem.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss')) from '$($outlookItem.SenderEmailAddress)'"
            Write-Debug   -Message "      Subject: '$($outlookItem.Subject)'"

            #Only update the progress bar ever 1000 Milliseconds, otherwise run time is > 100 slower
            if ($writeProgressStopWatch.Elapsed.TotalMilliseconds -ge 1000) {
                $writeProgressParameters = @{
                    Activity        = "Processing $($outlookStore.DisplayName)"
                    Status          = "Item $writeProgressCounter of $($outlookItems.Count)"
                    PercentComplete = (($writeProgressCounter/($outlookItems.Count)) * 100)
                }

                Write-Progress @writeProgressParameters
                $writeProgressStopWatch.Reset()
                $writeProgressStopWatch.Start()
            }

            $attachments = @($outlookItem.Attachments)
            Write-Debug -Message "    Attachment Count: $($attachments.Count)"
            try {
                foreach ($attachment in $attachments) {
                    Write-Debug -Message "      Attachment Name: '$($attachment.FileName)'"
                    $validAttachmentExtension = Test-StringAgainstFilterList -String $attachment.FileName -FilterList @(
                        '*.7z',
                        '*.gz',
                        '*.rar',
                        '*.tar',
                        '*.x',
                        '*.zip'
                    )

                    Write-Debug -Message "        Attachment uses a valid file extension: $($validAttachmentExtension.ToString())"

                    if ($validAttachmentExtension -eq $true) {
                        $attachmentDirectory = Join-Path -Path $mailboxScrapeDirectory -ChildPath $attachmentCounter
                        New-Item -Path $attachmentDirectory -ItemType Directory -Force | Out-Null

                        $attachmentFilePath = Join-Path -Path $attachmentDirectory -ChildPath $attachment.FileName
                        Write-Debug -Message "        Attachment download location: $attachmentFilePath"

                        $attachment.SaveAsFile($attachmentFilePath)

                        #When finished with the item, move it to the destination folder (i.e. completed items).
                        Write-Debug   -Message "      Moving item to '$($outlookDestinationFolder.FolderPath)'"
                        $outlookItem.Move($outlookDestinationFolder) | Out-Null

                    } else {
                        Write-Error "Not valid extension for DMARC report '$($attachment.FileName)'."
                    }

                    Remove-Variable -Name attachment
                    $attachmentCounter++
                }
            } catch {
                $attachmentCounter++
                Write-Error $_
            }
            $outlookItemCounter++
            $writeProgressCounter++

        } While ($outlookItemCounter -lt $ProcessingLimit -and $outlookItemCounter -lt ($outlookItems.Count - 1))

        #Output the reason why the processing has completed.
        switch ($outlookItemCounter) {
            {-not ($outlookItemCounter -lt $ProcessingLimit)}          {Write-Verbose -Message "End processing $($outlookStore.DisplayName)\$SourceFolder (ProcessingLimit of $ProcessingLimit reached)"; Break}
            {-not ($outlookItemCounter -lt ($outlookItems.Count - 1))} {Write-Verbose -Message "End processing $($outlookStore.DisplayName)\$SourceFolder (all $($outlookItems.Count) items processed)"}
        }
    }
}

