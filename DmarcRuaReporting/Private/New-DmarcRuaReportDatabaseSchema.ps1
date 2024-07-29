
Function New-DmarcRuaReportDatabaseSchema {
    [CmdletBinding()]
    Param (
        [System.String[]]
        $OutlookMailboxesToScrape,

        [System.String]
        $7zipExecutablePath = 'C:\Program Files\7-Zip\7z.exe'
    )

    $drrConfig = @{
        UnprocessedRawAttachment = @{
            _Comment     = "Contains all the raw files that were previously attached to emails sent to the DMARC RUA email address. These files are generally compressed and have a file extension of .gz or .zip and have yet to be added to the database (i.e. unprocessed)."
            DefaultValue = '.\Databases\UnprocessedRawAttachment'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Container'
                NewItemType    = 'Directory'
            }
        }

        UnprocessedXmlRuaReport  = @{
            _Comment     = "Contains all the xml files that were previously compressed and attached to emails sent to the DMARC RUA email address. These files must have a file extension of .xml and have yet to be added to the database (i.e. unprocessed)."
            DefaultValue = '.\Databases\UnprocessedXmlRuaReport'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Container'
                NewItemType    = 'Directory'
            }
        }

        ProcessedRawAttachment   = @{
            _Comment     = "Contains all the raw files that were previously attached to emails sent to the DMARC RUA email address. These files are generally compressed and have a file extension of .gz or .zip and have already been added to the database (i.e. processed)."
            DefaultValue = '.\Databases\ProcessedRawAttachment'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Container'
                NewItemType    = 'Directory'
            }
        }

        ProcessedXmlRuaReport    = @{
            _Comment     = "Contains all the xml files that were previously compressed and attached to emails sent to the DMARC RUA email address. These files must have a file extension of .xml and have already been added to the database (i.e. processed)."
            DefaultValue = '.\Databases\ProcessedXmlRuaReport'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Container'
                NewItemType    = 'Directory'
            }
        }

        DmarcReportDatabaseFiles = @{
            _Comment     = "Contains all the files that make up the DMARC RUA report database. These are stored in CSV files based on the date the emails were sent."
            DefaultValue = '.\Databases\DMARCReportDatabase'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Container'
                NewItemType    = 'Directory'
            }
        }

        IpToServerNameDB         = @{
            _Comment     = "A database cache of all the IP to Server Name mappings resolved using a DNS PTR lookup. Its is to reduce the frequency of which DNS queries are made."
            DefaultValue = '.\Databases\IP-to-ServerName-DB.xml'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Leaf'
                NewItemType    = 'File'
            }
        }

        ProcessedRawAttachmentDB = @{
            _Comment     = "A database containing a list of file hashes of all files in the ProcessedRawAttachment directory. Its purpose is to prevent duplicate data from being added to the database from files that have already been processed."
            DefaultValue = '.\Databases\ProcessedRawAttachment-DB.xml'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Leaf'
                NewItemType    = 'File'
            }
        }

        ProcessedXmlRuaReportDB  = @{
            _Comment     = "A database containing a list of file hashes of all files in the ProcessedXmlRuaReport directory. Its purpose is to prevent duplicate data from being added to the database from files that have already been processed."
            DefaultValue = '.\Databases\ProcessedXmlRuaReport-DB.xml'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Leaf'
                NewItemType    = 'File'
            }
        }

        ImportSourceDirectory        = @{
            _Comment     = "The input location from which to import DMARC RUA reports from. These files can either be in a compressed format (i.e. with file extension of .gz or .zip) or decompressed (i.e. with file extension of .xml). Multiple directories can be specified and files manually moved to one of these locations will be imported into the database."
            DefaultValue = '.\ImportData'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Container'
                NewItemType    = 'Directory'
            }
        }

        ReportOutputDirectory        = @{
            _Comment     = "The output location for all reports run against the DMARC RUA Reporting database."
            DefaultValue = '.\Reports'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Container'
                NewItemType    = 'Directory'
            }
        }

        ImportErrorDirectory         = @{
            _Comment     = "The location for all reports that failed to get processed and imported into the DMARC RUA Reporting database. Files can end up here if they contain an invalid or missing file extension."
            DefaultValue = '.\zz-ImportError'
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Container'
                NewItemType    = 'Directory'
            }
        }

        ApplicationExecutable_7zip   = @{
            _Comment     = "The file path of the 7-zip application. This must be installed to run this script as it is used to extract .xml files from the compressed file attachments."
            DefaultValue = $7zipExecutablePath
            ValueType    = @{
                IsPath         = $true
                TestPathType   = 'Leaf'
                NonReplaceable = $true
            }
        }

        OutlookMailboxesToScrape     = @{
            _Comment     = "The list of Outlook mailboxes to scrape for new emails containing DMARC RUA reports."
            DefaultValue = $OutlookMailboxesToScrape
            ValueType    = @{
                IsPath         = $false
            }
        }
    }

    Write-Output -InputObject $drrConfig
}
