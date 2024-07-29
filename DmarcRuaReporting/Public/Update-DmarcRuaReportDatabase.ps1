Function Update-DmarcRuaReportDatabase {
    [CmdletBinding()]
    Param (
        [ValidateScript({Test-Path -LiteralPath $_ -PathType Container})]
        [System.String]
        $Path = (Get-Item -Path .).FullName
    )

    $testDmarcRuaReportDatabase = Test-DmarcRuaReportDatabase -Path $Path
    if ($testDmarcRuaReportDatabase -eq $false) {
        throw "Could not find a valid DMARC RUA Report Database in the path '$Path'."
    }

    #Import all paths in the DevicePath format in order to get around the 256 character path limit in Windows Explorer.
    $drrConfigData = Get-DmarcRuaReportDatabase -PathFormat 'DevicePath'

    $rawEmailAttachmentAcceptedFileExtensions = @(
        '*.7z',
        '*.gz',
        '*.rar',
        '*.tar',
        '*.x',
        '*.zip'
    )


    <#
        Generate a file hash of each raw email attachment file in .\ImportData and check if it exists in the database Databases\ProcessedRawAttachment-DB.xml.
        If the hash is not in the database, then the file has never been seen before and is placed in .\Databases\UnprocessedRawAttachment.
        If the hash is already in the database, then delete the file as it's been processed before.
    #>
    Write-Verbose -Message @"
---- Step 1 of 4 ----
Importing raw email attachment files that haven't previously been processed based on their file hash.
  Import Path:     $($drrConfigData.ImportSourceDirectory)
  Export Path:     $($drrConfigData.UnprocessedRawAttachment)
  File Extensions: $([String]$rawEmailAttachmentAcceptedFileExtensions)

"@
    Add-FileToHashDatabase `
    -FileHashDatabase              $drrConfigData.ProcessedRawAttachmentDB `
    -ImportDirectoryPath           $drrConfigData.ImportSourceDirectory `
    -ExportDirectoryPath           $drrConfigData.UnprocessedRawAttachment `
    -ErrorsProcessingDirectoryPath $drrConfigData.ImportErrorDirectory `
    -AcceptedFileExtension         $rawEmailAttachmentAcceptedFileExtensions `
    -IgnoreFileExtension           '*.xml'


    <#
        Expand (Unzip) all the raw email attachments files in .\Databases\UnprocessedRawAttachment.
        Place the unzipped contents in .\ImportData.
        Place the original files in .\Databases\ProcessedRawAttachment
    #>
    Write-Verbose -Message @"
---- Step 2 of 4 ----
Expanding/unzipping raw email attachment files to extract the compressed .xml based report inside them.
  Import Path:              $($drrConfigData.UnprocessedRawAttachment)
  Export Path (.xml files): $($drrConfigData.ImportSourceDirectory)
  Backup of original files: $($drrConfigData.ProcessedRawAttachment)
  File Extensions:          $([String]$rawEmailAttachmentAcceptedFileExtensions)

"@
    Expand-7zipArchive `
    -ImportDirectoryPath           $drrConfigData.UnprocessedRawAttachment `
    -ExportDirectoryPath           $drrConfigData.ImportSourceDirectory `
    -CompletedDirectoryPath        $drrConfigData.ProcessedRawAttachment `
    -ErrorsProcessingDirectoryPath $drrConfigData.ImportErrorDirectory `
    -AcceptedFileExtension         $rawEmailAttachmentAcceptedFileExtensions `
    -SevenZipExecutablePath        $drrConfigData.ApplicationExecutable_7zip


    <#
        Generate a file hash of each raw .xml rua report file in .\ImportData and check if it exists in the database Databases\ProcessedXmlRuaReport-DB.xml.
        If the hash is not in the database, then the file has never been seen before and is placed in .\Databases\UnprocessedXmlRuaReport.
        If the hash is already in the database, then delete the file as it's been processed before.
    #>
    Write-Verbose -Message @"
---- Step 3 of 4 ----
Importing .xml DMARC RUA report files that haven't previously been processed based on their file hash.
    Import Path:     $($drrConfigData.ImportSourceDirectory)
    Export Path:     $($drrConfigData.UnprocessedXmlRuaReport)
    File Extensions: *.xml

"@
    Add-FileToHashDatabase `
    -FileHashDatabase              $drrConfigData.ProcessedXmlRuaReportDB `
    -ImportDirectoryPath           $drrConfigData.ImportSourceDirectory `
    -ExportDirectoryPath           $drrConfigData.UnprocessedXmlRuaReport `
    -ErrorsProcessingDirectoryPath $drrConfigData.ImportErrorDirectory `
    -AcceptedFileExtension         '*.xml' `
    -IgnoreFileExtension           @(
        '*.7z',
        '*.gz',
        '*.rar',
        '*.tar',
        '*.x',
        '*.zip'
    )


    <#
        Parse the XML based rua report files in .\Databases\UnprocessedXmlRuaReport.
        Process each record in the report and export each one to a CSV file .\Databases\DMARCReportDatabase\<DATE OF REPORT>\DMARC_Processed_Data_<DOMAIN NAME>.csv
        Move the original file to .\Databases\ProcessedXmlRuaReport.

    #>
    Write-Verbose -Message @"
---- Step 4 of 4 ----
Parsing .xml DMARC RUA report files and converting to CSV format.
    Import Path:              $($drrConfigData.UnprocessedXmlRuaReport)
    Export Path (.csv files): $($drrConfigData.DmarcReportDatabaseFiles)
    Backup of original files: $($drrConfigData.ProcessedXmlRuaReport)
    File Extensions:          *.xml

"@
    Import-DmarcRuaReport `
    -ImportDirectoryPath             $drrConfigData.UnprocessedXmlRuaReport `
    -ExportDirectoryPath             $drrConfigData.DmarcReportDatabaseFiles `
    -CompletedDirectoryPath          $drrConfigData.ProcessedXmlRuaReport `
    -ErrorsProcessingDirectoryPath   $drrConfigData.ImportErrorDirectory `
    -ProcessedXmlRuaReportDBFilePath $drrConfigData.ProcessedXmlRuaReportDB `
    -IPToServerNameDatabaseFilePath  $drrConfigData.IpToServerNameDB `
    -AcceptedFileExtension           '*.xml' `
    -MaxPageSize                     1000
}
