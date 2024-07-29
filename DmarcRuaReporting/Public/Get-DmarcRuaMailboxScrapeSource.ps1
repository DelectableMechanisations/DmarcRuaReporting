<#
    .SYNOPSIS
        Gets a list of mailboxes and determines if they could be used as a scrape source.

    .DESCRIPTION
        The Get-DmarcRuaMailboxScrapeSource function gets a list of mailboxes, or searches for mailboxes and determines if they're suitable for using as a scrape source.
        Once a mailbox has been identified, the Start-DmarcRuaMailboxScrape function can be used to scrape the mailbox for emails and download any attachments.

        Note: The Outlook application needs to be running before you can run this function.

    .PARAMETER DisplayName
        An optional Display Name filter that can be specified when searching for mailbox scrape sources.
        Defaults to '*'.

    .EXAMPLE
        Get-DmarcRuaMailboxScrapeSource

        Gets a list of mailboxes that the current user has access to and determines if they are suitable to use as a scrape source.

        DisplayName                                 IsValidMailboxScrapeSource FilePath
        -----------                                 -------------------------- --------
        my-shared@primarydomain.com                                       True C:\Users\dmuser\AppData\Local\Microsoft\Outl…
        DM@primarydomain.com                                              True C:\Users\dmuser\AppData\Local\Microsoft\Outl…
        rua@primarydomain.com                                             True C:\Users\dmuser\AppData\Local\Microsoft\Outl…
        Online Archive - DM@primarydomain.com                            False
        Public Folders - my-shared@primarydomain.com                     False
        Public Folders - DM@primarydomain.com                            False
        Public Folders - rua@primarydomain.com                           False

    .EXAMPLE
        Get-DmarcRuaMailboxScrapeSource -DisplayName *rua*

        Gets any mailboxes with a Display Name containing the string '*rua*'.

        DisplayName                                 IsValidMailboxScrapeSource FilePath
        -----------                                 -------------------------- --------
        rua@primarydomain.com                                             True C:\Users\dmuser\AppData\Local\Microsoft\Outl…
        Public Folders - rua@primarydomain.com                           False
#>
Function Get-DmarcRuaMailboxScrapeSource {
    [CmdletBinding()]
    Param (
        [System.String[]]
        $DisplayName = '*'
    )

    #Confirm Outlook is running.
    $testOutlookRunning = @(Get-Process -Name Outlook*)
    if ($testOutlookRunning.Count -eq 0) {
        throw "Microsoft Outlook is not running. Please open Outlook and re-run this command."
    }

    #Connect to the Outlook API.
    $outlook = New-Object -ComObject 'Outlook.Application' -ErrorAction 'Stop'

    #Loop through each of the filters in $DisplayName and compare them to the Display Names of the current user's Outlook stores.
    $mailboxScrapeSources = New-Object -TypeName System.Collections.ArrayList
    foreach ($displayNameFilter in $DisplayName) {
        $tempMailboxScrapeSources = @($outlook.GetNamespace('MAPI').Stores | Where-Object {$_.DisplayName -like $displayNameFilter})
        Write-Debug -Message "Found $($tempMailboxScrapeSources.Count) mailbox scrape sources with a DisplayName like '$displayNameFilter'."
        $tempMailboxScrapeSources | ForEach-Object {[VOID]$mailboxScrapeSources.Add($_)}
    }

    $mailboxScrapeSources = @($mailboxScrapeSources | Sort-Object -Property DisplayName -Unique)

    #Check each filtered mailbox store for its suitability as a scrape source.
    foreach ($mailboxScrapeSource in $mailboxScrapeSources) {
        Test-DmarcRuaMailboxScrapeSource -OutlookStore $mailboxScrapeSource | Write-Output
    }
}
