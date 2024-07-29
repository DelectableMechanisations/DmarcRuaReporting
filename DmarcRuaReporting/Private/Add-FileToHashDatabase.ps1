Function Add-FileToHashDatabase {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]
        [System.String]
        $FileHashDatabasePath,

        [Parameter(Mandatory)]
        [System.String]
        $ImportDirectoryPath,

        [Parameter(Mandatory)]
        [System.String]
        $ExportDirectoryPath,

        [Parameter(Mandatory)]
        [System.String]
        $ErrorsProcessingDirectoryPath,

        [System.String[]]
        $AcceptedFileExtension,

        [System.String[]]
        $IgnoreFileExtension,

        [System.String]
        $Algorithm = 'SHA1'
    )

    Begin {
        $writeProgressStopWatch  = [System.Diagnostics.Stopwatch]::StartNew()
        $writeProgressCounter    = 1
        $newFilesAddedToDatabase = 0

        Write-Debug -Message "Start processing using File Hash database '$FileHashDatabasePath'."

        if ((Get-Item -LiteralPath $FileHashDatabasePath).Length -eq 0) {
            Write-Debug -Message "  Total Keys: 0 (recreating)"
            $fileHashDatabase = @{}

        } else {
            $fileHashDatabase = Import-CliXml -LiteralPath $FileHashDatabasePath -ErrorAction Stop
            Write-Debug -Message "  Total Keys: $($fileHashDatabase.Keys.Count)"
        }

        #Place a lock on the file to prevent other processes from accessing it.
        $fileLock = [System.IO.File]::Open($FileHashDatabasePath, 'Open', 'ReadWrite', 'None')

        $importFiles = @(Get-ChildItem -LiteralPath $ImportDirectoryPath -File -Recurse)
    }

    Process {
        foreach ($file in $importFiles) {
            Write-Debug -Message "  File: $($file.FullName)"

            $parentDirectory = $file.DirectoryName

            #Only update the progress bar ever 100 Milliseconds, otherwise run time is > 100 slower
            if ($writeProgressStopWatch.Elapsed.TotalMilliseconds -ge 100) {
                $writeProgressParameters = @{
                    Activity        = 'Processing Files...'
                    Status          = "File $writeProgressCounter of $($importFiles.Count)"
                    PercentComplete = ($writeProgressCounter/$importFiles.Count*100)
                }

                Write-Progress @writeProgressParameters
                $writeProgressStopWatch.Reset()
                $writeProgressStopWatch.Start()
            }

            #Confirm the file has an accepted file extension.
            if (Test-StringAgainstFilterList -String $file.Name -FilterList $AcceptedFileExtension) {
                $fileHash = Get-FileHash -LiteralPath $file.FullName -Algorithm $Algorithm

                #If the hash of the file already exists in the database then remove it.
                if (Test-FileHashInDatabase -FileHashDatabase $fileHashDatabase -FileHash $fileHash.Hash) {
                    Write-Debug -Message "    Hash already exists in the database. Removing file"
                    Remove-Item -LiteralPath $file.FullName -Force

                #If the hash of the file isn't currently in the database then rename it, move it to the $ExportDirectoryPath and add to the database.
                } else {
                    #Remove the existing file hash if present before renaming the file.
                    if ($file.Name -like 'UniqueHASH-*') {
                        $nameWithoutExistingHash = $file.Name.SubString($file.Name.IndexOf('_') + 1)
                        $newFilePath = "$ExportDirectoryPath\UniqueHASH-$($fileHash.Hash)_$($nameWithoutExistingHash)"

                    } else {
                        $newFilePath = "$ExportDirectoryPath\UniqueHASH-$($fileHash.Hash)_$($file.Name)"
                    }

                    Write-Debug -Message "    Moving file to '$newFilePath'."
                    Move-Item -LiteralPath $file.FullName -Destination $newFilePath -Force -ErrorAction Stop
                    $fileHashDatabase.Add($fileHash.Hash, $file.Name)
                    $newFilesAddedToDatabase++
                }

            #If the file's extension appears in the ignore list then do nothing.
            } elseif (Test-StringAgainstFilterList -String $file.Name -FilterList $IgnoreFileExtension) {
                Write-Debug -Message '    Ignored due to its file extension (it will be processed separately).'

            #If the file does not have an accepted file extension then move to $ErrorsProcessingDirectoryPath.
            } else {
                Write-Debug -Message "    Invalid file extension. This will be moved to '$ErrorsProcessingDirectoryPath'."
                Move-Item -LiteralPath $file.FullName -Destination "$ErrorsProcessingDirectoryPath\$(New-Guid)_$($file.Name)" -Force -ErrorAction Stop
            }

            #Remove the file's parent directory if it is empty.
            Remove-EmptyDirectory -LiteralPath $parentDirectory -IncludeRootPath

            $writeProgressCounter++
        }
    }

    End {
        Write-Debug -Message "$newFilesAddedToDatabase file(s) have been added to the database."

        #Remove the file lock and export the data to the database.
        $fileLock.Close()
        $fileHashDatabase | Export-CliXml -LiteralPath $FileHashDatabasePath -ErrorAction Stop

        #Remove all empty directories from the import path.
        Remove-EmptyDirectory -LiteralPath $ImportDirectoryPath
    }
}
