<#
    .SYNOPSIS
        Gets the details of a DMARC RUA Report Database.

    .DESCRIPTION
        Use this function to return the metadata details of a DMARC RUA Report Database stored in the 'config.json' file.
        Metadata returned by this function includes database paths, the path to the 7-Zip executable and mailbox scrape sources.
        It is primarily used by other functions, negating the need for these paths to be hard coded into each function.

    .PARAMETER Path
        Specifies the directory root path to the DMARC RUA Report Database.
        Defaults to . (the current location).

    .PARAMETER PathFormat
        Specifies the path format to use when returning the file and directory paths used by the DMARC RUA Report Database.
            AbsolutePath - Return paths in their full, absolute form (e.g. C:\Temp\DmarcRuaReportDatabase\Databases\DMARCReportDatabase).
            DevicePath   - Return paths in the device path format that supports paths with names longer than 255 characters (e.g. \\?\C:\Temp\DmarcRuaReportDatabase\Databases\DMARCReportDatabase).
            RelativePath - Return paths relative to the directory root of the the DMARC RUA Report Database. This is the default value.

    .EXAMPLE
        Get-DmarcRuaReportDatabase

        Gets the details of the DMARC RUA Report Database in the current path.

        ApplicationExecutable_7zip : C:\Program Files\7-Zip\7z.exe
        DmarcReportDatabaseFiles   : .\Databases\DMARCReportDatabase
        ImportErrorDirectory       : .\zz-ImportError
        ImportSourceDirectory      : .\ImportData
        IpToServerNameDB           : .\Databases\IP-to-ServerName-DB.xml
        OutlookMailboxesToScrape   : {rua@primarydomain.com}
        ProcessedRawAttachment     : .\Databases\ProcessedRawAttachment
        ProcessedRawAttachmentDB   : .\Databases\ProcessedRawAttachment-DB.xml
        ProcessedXmlRuaReport      : .\Databases\ProcessedXmlRuaReport
        ProcessedXmlRuaReportDB    : .\Databases\ProcessedXmlRuaReport-DB.xml
        ReportOutputDirectory      : .\Reports
        UnprocessedRawAttachment   : .\Databases\UnprocessedRawAttachment
        UnprocessedXmlRuaReport    : .\Databases\UnprocessedXmlRuaReport

    .EXAMPLE
        Get-DmarcRuaReportDatabase -PathFormat AbsolutePath

        Gets the details of the DMARC RUA Report Database in the current path, and returns the file and directory paths as absolute paths.

        ApplicationExecutable_7zip : C:\Program Files\7-Zip\7z.exe
        DmarcReportDatabaseFiles   : C:\Temp\DmarcRuaReportDatabase\Databases\DMARCReportDatabase
        ImportErrorDirectory       : C:\Temp\DmarcRuaReportDatabase\zz-ImportError
        ImportSourceDirectory      : C:\Temp\DmarcRuaReportDatabase\ImportData
        IpToServerNameDB           : C:\Temp\DmarcRuaReportDatabase\Databases\IP-to-ServerName-DB.xml
        OutlookMailboxesToScrape   : {rua@primarydomain.com}
        ProcessedRawAttachment     : C:\Temp\DmarcRuaReportDatabase\Databases\ProcessedRawAttachment
        ProcessedRawAttachmentDB   : C:\Temp\DmarcRuaReportDatabase\Databases\ProcessedRawAttachment-DB.xml
        ProcessedXmlRuaReport      : C:\Temp\DmarcRuaReportDatabase\Databases\ProcessedXmlRuaReport
        ProcessedXmlRuaReportDB    : C:\Temp\DmarcRuaReportDatabase\Databases\ProcessedXmlRuaReport-DB.xml
        ReportOutputDirectory      : C:\Temp\DmarcRuaReportDatabase\Reports
        UnprocessedRawAttachment   : C:\Temp\DmarcRuaReportDatabase\Databases\UnprocessedRawAttachment
        UnprocessedXmlRuaReport    : C:\Temp\DmarcRuaReportDatabase\Databases\UnprocessedXmlRuaReport

    .EXAMPLE
        Get-DmarcRuaReportDatabase -Path 'C:\Temp\DmarcRuaReportDatabase' -PathFormat DevicePath

        Gets the details of the DMARC RUA Report Database in the path 'C:\Temp\DmarcRuaReportDatabase' , and returns the file and directory paths as device paths.

        ApplicationExecutable_7zip : C:\Program Files\7-Zip\7z.exe
        DmarcReportDatabaseFiles   : \\?\C:\Temp\DmarcRuaReportDatabase\Databases\DMARCReportDatabase
        ImportErrorDirectory       : \\?\C:\Temp\DmarcRuaReportDatabase\zz-ImportError
        ImportSourceDirectory      : \\?\C:\Temp\DmarcRuaReportDatabase\ImportData
        IpToServerNameDB           : \\?\C:\Temp\DmarcRuaReportDatabase\Databases\IP-to-ServerName-DB.xml
        OutlookMailboxesToScrape   : {rua@primarydomain.com}
        ProcessedRawAttachment     : \\?\C:\Temp\DmarcRuaReportDatabase\Databases\ProcessedRawAttachment
        ProcessedRawAttachmentDB   : \\?\C:\Temp\DmarcRuaReportDatabase\Databases\ProcessedRawAttachment-DB.xml
        ProcessedXmlRuaReport      : \\?\C:\Temp\DmarcRuaReportDatabase\Databases\ProcessedXmlRuaReport
        ProcessedXmlRuaReportDB    : \\?\C:\Temp\DmarcRuaReportDatabase\Databases\ProcessedXmlRuaReport-DB.xml
        ReportOutputDirectory      : \\?\C:\Temp\DmarcRuaReportDatabase\Reports
        UnprocessedRawAttachment   : \\?\C:\Temp\DmarcRuaReportDatabase\Databases\UnprocessedRawAttachment
        UnprocessedXmlRuaReport    : \\?\C:\Temp\DmarcRuaReportDatabase\Databases\UnprocessedXmlRuaReport
#>
Function Get-DmarcRuaReportDatabase {
    [CmdletBinding()]
    Param (
        [ValidateScript({Test-Path -Path $_ -PathType Container })]
        [System.String]
        $Path = (Get-Item -Path .).FullName,

        [ValidateSet('AbsolutePath', 'DevicePath', 'RelativePath')]
        [System.String]
        $PathFormat = 'RelativePath'
    )

    #Confirm that a 'config.json' file exists in the root of $Path.
    $drrConfigFilePath = Join-Path -Path $Path -ChildPath 'config.json' -ErrorAction 'SilentlyContinue'

    if (Test-Path -Path $drrConfigFilePath) {
        $drrConfig = Get-Content -Path $drrConfigFilePath | ConvertFrom-Json -Depth 10 -AsHashtable

    } else {
        throw "Unable to find '$Path\config.json'. Please confirm you are in the current path or run the New-DmarcRuaReportDatabase function to create it."
    }

    $drrConfigItems = $drrConfig.Keys | ForEach-Object {$_.ToString()} | Sort-Object
    $drrConfigData = New-Object -TypeName psobject
    foreach ($drrConfigItem in $drrConfigItems) {
        Write-Debug -Message "Importing property '$drrConfigItem'."

        $drrConfigItemValue = ($drrConfig[$drrConfigItem]).Value

        #If the property is a relative path then (if required) convert this as per the PathFormat variable.
        if ($drrConfigItemValue -like '.\*') {
            $absolutePath = [System.IO.Path]::GetFullPath($drrConfigItemValue, $Path)
            switch ($PathFormat) {
                #Absolute paths begin with a drive letter (e.g. C:\Temp\DmarcRuaReportDatabase)
                'AbsolutePath' {$drrConfigItemValue = $absolutePath}

                #Device paths aren't subject to the 255 character path limits (e.g. \\?\C:\Temp\DmarcRuaReportDatabase)
                'DevicePath'   {$drrConfigItemValue = "\\?\$absolutePath"}

                #Relative paths are relative to the current location in the file system (e.g. .\DmarcRuaReportDatabase)
                'RelativePath' {$drrConfigItemValue = $drrConfigItemValue}
            }
        }

        $drrConfigData | Add-Member -MemberType NoteProperty -Name $drrConfigItem -Value $drrConfigItemValue
    }

    Write-Output -InputObject $drrConfigData
}
