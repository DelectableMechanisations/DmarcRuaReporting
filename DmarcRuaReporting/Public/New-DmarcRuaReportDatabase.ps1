<#
    .SYNOPSIS
        Creates a new DMARC RUA Report Database.

    .DESCRIPTION
        The New-DmarcRuaReportDatabase function is used to create the folder structure, metadata and database files used for collecting, storing and parsing DMARC RUA Reports.
        Creating a dedicated folder structure like this allows a much simpler set of functions to be used to keep the database up to date.

    .PARAMETER Name
        The name of the DMARC RUA Report Database root folder name.
        Defaults to 'DmarcRuaReportDatabase'.

    .PARAMETER OutlookMailboxesToScrape
        Specifies one or more Outlook mailboxes that email attachments should be scraped from.
        Use the Get-DmarcRuaMailboxScrapeSource function to determine valid mailboxes.

    .PARAMETER Path
        Specifies the directory root path to the DMARC RUA Report Database.
        Defaults to . (the current location).

    .PARAMETER Force
        Forces this function to overwrite an existing DMARC RUA Report Database.

    .PARAMETER 7zipExecutablePath
        The path to the 7-Zip executable.
        Defaults to 'C:\Program Files\7-Zip\7z.exe'.

    .EXAMPLE
        PS C:\Temp> New-DmarcRuaReportDatabase -OutlookMailboxesToScrape 'rua@primarydomain.com'
        PS C:\Temp\DmarcRuaReportDatabase>

        Creates a new DMARC RUA Report Database in the current path (C:\Temp) and specifying the mailbox 'rua@primarydomain.com' as a scrape source.
        Once the database has been created, the CLI changes into that directory.

    .EXAMPLE
        PS C:\> New-DmarcRuaReportDatabase `
        -Name                     'MyDmarcRuaDB' `
        -OutlookMailboxesToScrape 'rua@primarydomain.com' `
        -Path                     'C:\Work' `
        -7zipExecutablePath       'C:\Work\7z.exe'

        PS C:\Work\DmarcRuaReportDatabase>

        Creates a new DMARC RUA Report Database called 'MyDmarcRuaDB' in the path 'C:\Work'.
        Also specifies a custom 7-Zip executable path and that the mailbox 'rua@primarydomain.com' should be used as the scrape source.
#>
Function New-DmarcRuaReportDatabase {
    [CmdletBinding()]
    Param (
        [System.String]
        $Name = 'DmarcRuaReportDatabase',

        [System.String[]]
        $OutlookMailboxesToScrape,

        [ValidateScript({Test-Path -Path $_ -PathType Container -IsValid})]
        [System.String]
        $Path = (Get-Item -Path .).FullName,

        [Switch]
        $Force,

        [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
        [System.String]
        $7zipExecutablePath = 'C:\Program Files\7-Zip\7z.exe'
    )

    $drrDirectoryPath = Join-Path -Path $Path -ChildPath $Name

    if (Test-Path -Path $drrDirectoryPath) {
        if ($Force -eq $true) {
            Remove-Item -Path $drrDirectoryPath -Force -Recurse -ErrorAction Stop

        } else {
            throw "Path '$drrDirectoryPath' already exists."
        }
    }

    New-Item -Path $drrDirectoryPath -ItemType Directory -ErrorAction Stop | Out-Null
    Set-Location -Path $drrDirectoryPath

    $testDmarcRuaReportDatabaseParams = (New-Object -TypeName System.Collections.Hashtable -ArgumentList $PSBoundParameters).Clone()
    Remove-FromPSBoundParametersUsingHashtable -PSBoundParametersIn $testDmarcRuaReportDatabaseParams -ParamsToKeep @('7zipExecutablePath', 'OutlookMailboxesToScrape')


    Test-DmarcRuaReportDatabase -Path $drrDirectoryPath -AddMissing @testDmarcRuaReportDatabaseParams | Out-Null
}
