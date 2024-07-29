<#
    .SYNOPSIS
        Removes all empty directory from within a directory path.
        Any directories containing files or other directories will be skipped.
#>
Function Remove-EmptyDirectory {
    [CmdletBinding()]
    Param (
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
        [System.String]
        $LiteralPath,

        [Switch]
        $IncludeRootPath
    )

    Write-Debug -Message "Scanning '$LiteralPath' for empty directories to remove."
    $directories = @(Get-ChildItem -LiteralPath $LiteralPath -Directory -Recurse | Select-Object -ExpandProperty FullName)
    [System.Array]::Reverse($directories)

    if ($IncludeRootPath) {
        $directories += $LiteralPath
        Write-Debug -Message "  Parameter '-IncludeRootPath' has been specified. Root path '$LiteralPath' will also be removed if it is empty."
    }

    foreach ($directory in $directories) {
        Write-Debug -Message "  Checking subdirectory '$directory' for items."
        $contentsTest = @(Get-ChildItem -LiteralPath $directory -Force)
        if ($contentsTest.Count -eq 0) {
            Write-Debug -Message "    No items found. Removing."
            Remove-Item -LiteralPath $directory

        } else {
            Write-Debug -Message "    Found $($contentsTest.Count) items. Skipping removal."
        }
    }
}
