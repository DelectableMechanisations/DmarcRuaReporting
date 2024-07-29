Function Expand-7zipArchive {
    [CmdletBinding()]
    Param (
        #Default:    .\Databases\UnprocessedRawAttachment
        [Parameter(Mandatory)]
        [System.String]
        $ImportDirectoryPath,

        #Default:    .\ImportData
        [Parameter(Mandatory)]
        [System.String]
        $ExportDirectoryPath,

        #Default:    .\Databases\ProcessedRawAttachment
        [Parameter(Mandatory)]
        [System.String]
        $CompletedDirectoryPath,

        #Default:    .\zz-ImportError
        [Parameter(Mandatory)]
        [System.String]
        $ErrorsProcessingDirectoryPath,

        [System.String[]]
        $AcceptedFileExtension,

        [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]
        [System.String]
        $SevenZipExecutablePath = 'C:\Program Files\7-Zip\7z.exe'
    )

    Begin {
        $writeProgressStopWatch  = [System.Diagnostics.Stopwatch]::StartNew()
        $writeProgressCounter    = 1

        #Create a temporary working directory that will store all the expanded archives.
        $date = Get-Date -Format 'yyyy-MM-dd HHmmss'
        $tempWorkingDirectoryRoot = "$ExportDirectoryPath\ExpandArchive_$($date)"
        Write-Debug -Message "Creating a temporary working directory '$tempWorkingDirectoryRoot'."
        New-Item -Path $tempWorkingDirectoryRoot -ItemType Directory -Force | Out-Null

        #Create a temporary directory to store any files that generate an error when expanding.
        $tempErrorsProcessingDirectoryRoot = "$ErrorsProcessingDirectoryPath\ExpandArchive_$($date)"
        Write-Debug -Message "Creating a temporary errors directory root '$tempErrorsProcessingDirectoryRoot'."
        New-Item -Path $tempErrorsProcessingDirectoryRoot -ItemType Directory -Force | Out-Null

        $importFiles = @(Get-ChildItem -LiteralPath $ImportDirectoryPath -File -Recurse)
    }

    Process {
        foreach ($file in $importFiles) {
            Write-Debug -Message "  File: $($file.FullName)"

            #Only update the progress bar ever 100 Milliseconds, otherwise run time is > 100 slower
            if ($writeProgressStopWatch.Elapsed.TotalMilliseconds -ge 100) {
                $writeProgressParameters = @{
                    Activity        = 'Unzipping files...'
                    Status          = "File $writeProgressCounter of $($importFiles.Count)"
                    PercentComplete = ($writeProgressCounter/$importFiles.Count*100)
                }

                Write-Progress @writeProgressParameters
                $writeProgressStopWatch.Reset()
                $writeProgressStopWatch.Start()
            }

            #Confirm the file has an accepted file extension.
            if (Test-StringAgainstFilterList -String $file.Name -FilterList $AcceptedFileExtension) {

                #Create temporary directory for each expanded archive file.
                $tempWorkingDirectory = "$tempWorkingDirectoryRoot\EA_$($file.Name)"
                New-Item -Path $tempWorkingDirectory -ItemType Directory -Force | Out-Null

                $7zipResult = & $SevenZipExecutablePath x $file.FullName -o"$tempWorkingDirectory" -bse1

                #If no issues when expanding the archive, move the archive into "$CompletedDirectoryPath".
                if (
                    [String]$7zipResult -match 'Everything is Ok' -or `
                    [String]$7zipResult -match 'There are some data after the end of the payload data'
                ) {
                    Write-Debug -Message "    Successfully expanded archive file into '$tempWorkingDirectory'."
                    Write-Debug -Message "    Moving original file into '$CompletedDirectoryPath'."
                    Move-Item -LiteralPath $file.FullName -Destination $CompletedDirectoryPath

                #If there are issues when expanding the archive then move to "$tempErrorsProcessingDirectoryRoot".
                } else {
                    Write-Debug -Message "    Failed expanding archive file into '$tempWorkingDirectory'."
                    Write-Debug -Message "    Moving original file into '$tempErrorsProcessingDirectoryRoot'."
                    Move-Item -LiteralPath $file.FullName -Destination $tempErrorsProcessingDirectoryRoot
                }

            #If the file does not have an accepted file extension then move to "$tempErrorsProcessingDirectoryRoot".
            } else {
                Write-Debug -Message "    File has an invalid extension."
                Write-Debug -Message "    Moving original file into '$tempErrorsProcessingDirectoryRoot'."
                Move-Item -LiteralPath $file.FullName -Destination $tempErrorsProcessingDirectoryRoot
            }

            #Remove the temp working directory if it is empty.
            Remove-EmptyDirectory -LiteralPath $tempWorkingDirectory -IncludeRootPath

            $writeProgressCounter++
        }
    }

    End {
        #Remove the temporary error processing directory if it is empty.
        Remove-EmptyDirectory -LiteralPath $tempErrorsProcessingDirectoryRoot -IncludeRoot
    }
}
