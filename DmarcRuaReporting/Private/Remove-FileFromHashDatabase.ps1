Function Remove-FileFromHashDatabase {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]
        [System.String]
        $FileHashDatabasePath,

        [Parameter(Mandatory)]
        [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})]
        [System.String]
        $FilePath,

        [System.String]
        $Algorithm = 'SHA1'
    )

    Begin {
        Write-Debug -Message "Start processing using File Hash database '$FileHashDatabasePath'."

        if ((Get-Item -LiteralPath $FileHashDatabasePath).Length -eq 0) {
            Write-Debug   -Message "  Total Keys: 0 (recreating)"
            $fileHashDatabase = @{}

        } else {
            $fileHashDatabase = Import-CliXml -LiteralPath $FileHashDatabasePath -ErrorAction Stop
            Write-Debug   -Message "  Total Keys: $($fileHashDatabase.Keys.Count)"
        }

        #Place a lock on the file to prevent other processes from accessing it.
        $fileLock = [System.IO.File]::Open($FileHashDatabasePath, 'Open', 'ReadWrite', 'None')
    }

    Process {
        Write-Debug   -Message "  File: $($FilePath)"
        $fileHash = Get-FileHash -LiteralPath $FilePath -Algorithm $Algorithm

        #If the hash of the file exists in the database then remove its entry from the database.
        if (Test-FileHashInDatabase -FileHashDatabase $fileHashDatabase -FileHash $fileHash.Hash) {
            Write-Debug -Message "    Removing hash '$($fileHash.Hash)' from the database."
            $fileHashDatabase.Remove($fileHash.Hash)
        }
    }

    End {
        Write-Debug -Message "Finished processing using File Hash database '$FileHashDatabasePath'."

        #Remove the file lock and export the data to the database.
        $fileLock.Close()
        $fileHashDatabase | Export-CliXml -LiteralPath $FileHashDatabasePath -ErrorAction Stop
    }
}
