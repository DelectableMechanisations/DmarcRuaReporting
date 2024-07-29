Function Test-DmarcRuaMailboxScrapeSource {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [System.Object]
        $OutlookStore
    )

    Write-Debug -Message "Processing Outlook Mailbox Store '$($OutlookStore.DisplayName)'"
    $isValidMailboxScrapeSource = $true

    #Confirm Cached Exchange Mode is enabled for the mailbox.
    Write-Debug -Message "    IsCachedExchange = $($OutlookStore.IsCachedExchange.ToString())"
    if ($OutlookStore.IsCachedExchange -eq $false) {
        $isValidMailboxScrapeSource = $false
    }

    #Output the results of the test.
    Write-Output -InputObject ([PSCustomObject]@{
        DisplayName                = $OutlookStore.DisplayName
        IsValidMailboxScrapeSource = $isValidMailboxScrapeSource
        FilePath                   = $OutlookStore.FilePath
    })
}
