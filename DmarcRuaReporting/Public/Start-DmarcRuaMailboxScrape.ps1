<#
    .SYNOPSIS
        Downloads DMARC RUA reports sent to an Outlook mailbox.

    .DESCRIPTION
        The Start-DmarcRuaMailboxScrape function is used to scrape all DMARC RUA reports attached to emails that have been sent to an Outlook mailbox.
        High level actions are as follows:
            1) Connect to the Outlook mailbox(s) specified in the DMARC RUA report database or -DisplayName parameter.
            2) Enumerate through each email in the Outlook folder specified in the -SourceFolder parameter.
            3) If the email contains an attachment, confirm it matches a valid DMARC RUA report file extension (.7z .gz .rar .tar .x .zip).
            4) If the attachment does have a valid file extension then download it to a temporary 'MailboxScrape' folder in the DMARC RUA report database.
            5) Move the successfully processed email to the Outlook folder specified in the -DestinationFolder parameter.
            6) Any unprocssed emails are left in the source folder for the user to manually resolve.

        This function will sometimes fail to process one or more emails.
        Here are some tips to follow if this occurs:
            - Review the email and see if it contains an attachment.
              If it does, download the attachment and unzip its contents into the 'ImportData' directory.
              Chances are, the unzipped files don't have an '.xml' extension but are otherwise ok.
              Just rename all these files with an '.xml' extension and so long as the file is in the 'ImportData' directory, it will get picked up later on by the Update-DmarcRuaReportDatabase function.

            - Delete any spam emails in the Inbox folder or move them to another Outlook folder.
              If you don't do this, the function will display an error message each time it tries to process them.

            - If the function is failing on a large number of emails, it's recommend that you blow away your entire cached Outlook profile and re-download everything from scratch from Exchange Online.
    
    .PARAMETER Path
        Specifies the directory root path to the DMARC RUA Report Database.
        Defaults to . (the current location).

    .PARAMETER DisplayName
        An optional Display Name filter that can be specified when searching for mailbox scrape sources.
        Defaults to whatever is specified in the config.json file.

    .PARAMETER SourceFolder
        Specifies the source Outlook folder to scrape.
        Defaults to the 'Inbox' folder.

    .PARAMETER DestinationFolder
        Specifies the destination Outlook folder to move emails to once they have been processed.
        Defaults a custom mailbox root folder called '_DmarcProcessed'.

        Note: This parameter doesn't support specifying a subfolder (e.g. Inbox/_DmarcProcessed).

    .PARAMETER ProcessingLimit
        The maximum number of emails in Outlook to process in a single execution of this function.
        Defaults to 5000 and was introduced because Outlook tended to crash at anything higher than this on my 4 x CPU core, 8GB memory laptop.

    .EXAMPLE
        Start-DmarcRuaMailboxScrape

        Starts downloading all DMARC aggregate reports from the mailbox(s) specified by the DMARC RUA report database in the current path.
#>
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

