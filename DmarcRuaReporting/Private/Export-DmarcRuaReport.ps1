Function Export-DmarcRuaReport {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [System.Array]
        $DmarcRuaReport,

        [Parameter(Mandatory)]
        [System.String]
        $ExportDirectoryPath
    )

    Write-Debug -Message "Start processing $($DmarcRuaReport.Count) report items."
    $parentFolderNameGroups = $DmarcRuaReport | Sort-Object -Property 'DestinationReport_ParentFolderName' | Group-Object -Property 'DestinationReport_ParentFolderName'

    foreach ($parentFolderNameGroup in $parentFolderNameGroups) {
        $parentFolderPath = Join-Path -Path $ExportDirectoryPath -ChildPath $parentFolderNameGroup.Name

        if (-not (Test-Path -Path $parentFolderPath)) {
            Write-Debug -Message "    Creating new parent folder '$parentFolderPath'."
            New-Item -Path $parentFolderPath -ItemType Directory -Force | Out-Null
        }

        $fileNameGroups = $parentFolderNameGroup.Group | Sort-Object -Property 'DestinationReport_FileName' | Group-Object -Property 'DestinationReport_FileName'

        foreach ($fileNameGroup in $fileNameGroups) {
            $filePath = Join-Path -Path $parentFolderPath -ChildPath $fileNameGroup.Name
            Write-Debug -Message "    Outputting $($fileNameGroup.Group.Count) items to '$filePath'."
            $fileNameGroup.Group | Export-Csv -LiteralPath $filePath -NoTypeInformation -Append
        }
    }
}
