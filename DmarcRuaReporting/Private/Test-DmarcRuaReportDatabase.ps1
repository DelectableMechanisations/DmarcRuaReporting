#Confirm a DMARC RUA Report Database is present.
#Confirm a config file is present.
#Confirm all the folders in the config file are present.


Function Test-DmarcRuaReportDatabase {
    [CmdletBinding()]
    Param (
        [ValidateScript({ Test-Path -Path $_ -PathType Container })]
        [System.String]
        $Path = (Get-Item -Path .).FullName,

        [Switch]
        $AddMissing,

        [System.String[]]
        $OutlookMailboxesToScrape,

        [System.String]
        $7zipExecutablePath = 'C:\Program Files\7-Zip\7z.exe'
    )

    #Confirm that a 'config.json' file exists in the root of $Path.
    $drrConfigFilePath = Join-Path -Path $Path -ChildPath 'config.json' -ErrorAction 'SilentlyContinue'

    if (Test-Path -Path $drrConfigFilePath) {
        $drrConfig = Get-Content -Path $drrConfigFilePath | ConvertFrom-Json -Depth 10 -AsHashtable

    } elseif ($AddMissing -eq $true) {
        Write-Debug -Message "Could not find config file '$drrConfigFilePath'. As the '-AddMissing' parameter has been specified this will be recreated."
        $drrConfig = @{}

    } else {
        throw "Could not find config file '$drrConfigFilePath'."
    }


    #Create a new database schema. This is used to compare it to the existing database file and directory structure.
    $newDmarcRuaReportDatabaseSchemaParams = (New-Object -TypeName System.Collections.Hashtable -ArgumentList $PSBoundParameters).Clone()
    Remove-FromPSBoundParametersUsingHashtable -PSBoundParametersIn $newDmarcRuaReportDatabaseSchemaParams -ParamsToKeep @('7zipExecutablePath', 'OutlookMailboxesToScrape')

    $drrDatabaseSchema = New-DmarcRuaReportDatabaseSchema @newDmarcRuaReportDatabaseSchemaParams
    $drrDatabaseSchemaItems = New-Object -TypeName System.Collections.ArrayList -ArgumentList $drrDatabaseSchema.Keys


    #Loop through each item in the database schema.
    foreach ($drrDatabaseSchemaItem in $drrDatabaseSchemaItems) {
        Write-Debug -Message "Start processing database schema item $drrDatabaseSchemaItem."
        $currentSchemaItem = $drrDatabaseSchema[$drrDatabaseSchemaItem]


        #Confirm the config file contains the database schema item.
        if ($drrConfig.ContainsKey($drrDatabaseSchemaItem)) {
            Write-Debug -Message "    Schema item '$drrDatabaseSchemaItem' is present."

            #Confirm the database schema item contains the property 'Value'.
            $drrConfigKey = $drrConfig[$drrDatabaseSchemaItem]
            if ($drrConfigKey.ContainsKey('Value')) {
                Write-Debug -Message "    Schema item property 'Value' is present."

            } elseif ($AddMissing -eq $true) {
                Write-Debug -Message "    Schema item property 'Value' is missing. It will be recreated using the default value."
                $drrConfigKey['Value'] = $currentSchemaItem.DefaultValue

            } else {
                Write-Debug -Message "    Schema item property 'Value' is missing."
                throw "DmarcRuaReport config file '$drrConfigFilePath' is corrupted."
            }


            #If the database schema item is missing and the $AddMissing switch has been specified then add them using its default value.
        } elseif ($AddMissing -eq $true) {
            Write-Debug -Message "    Schema item $drrDatabaseSchemaItem is missing. It will be recreated using the default value."
            $drrConfig[$drrDatabaseSchemaItem] = @{
                Value    = $currentSchemaItem.DefaultValue
                _Comment = $currentSchemaItem._Comment
            }
            $drrConfigKey = $drrConfig[$drrDatabaseSchemaItem]


            #If the database schema item is missing and the $AddMissing switch has NOT been specified then terminate.
        } else {
            Write-Debug -Message "    Schema item $drrDatabaseSchemaItem is missing."
            throw "DmarcRuaReport config file '$drrConfigFilePath' is corrupted."
        }


        #Perform additional checks if the database schema item is a path.
        if ($currentSchemaItem.ValueType.IsPath -eq $true) {

            #Terminate if the path is non-replaceable (i.e. an application that needs to be installed).
            if ($currentSchemaItem.ValueType.NonReplaceable -eq $true) {
                if (Test-Path -Path $drrConfigKey.Value) {
                    Write-Debug -Message "    Path '$($drrConfigKey.Value)' is present."

                } else {
                    Write-Debug -Message "    Path '$($drrConfigKey.Value)' is missing."
                    throw "Mandatory path '$drrConfigKey.Value' is missing."
                }

            } else {

                #If the path is valid and of the correct type then do nothing.
                if (Test-Path -Path $drrConfigKey.Value -PathType $currentSchemaItem.ValueType.TestPathType) {
                    Write-Debug -Message "    Path '$($drrConfigKey.Value)' is present."

                    #If the path is valid but is the wrong type then terminate.
                } elseif (Test-Path -Path $drrConfigKey.Value) {
                    Write-Debug -Message "    Path '$($drrConfigKey.Value)' is present."
                    throw "Path '$($drrConfigKey.Value)' is an invalid path type. It should be type '$($currentSchemaItem.ValueType.PathType)'."

                    #If the path is missing but '-AddMissing' has been specified then recreate it.
                } elseif ($AddMissing -eq $true) {
                    Write-Debug -Message "    Path '$($drrConfigKey.Value)' is missing. It will be recreated."
                    New-Item -Path $drrConfigKey.Value -ItemType $currentSchemaItem.ValueType.NewItemType -Force | Out-Null

                    #If the path is missing but '-AddMissing' has NOT been specified then terminate.
                } else {
                    Write-Debug -Message "    Path '$($drrConfigKey.Value)' is missing."
                    throw "Path '$($drrConfigKey.Value)' is missing."
                }
            }
        }

        Write-Debug -Message "End processing database schema item $drrDatabaseSchemaItem."
    }

    #If the path is missing but '-AddMissing' has been specified then re-export the config file.
    if ($AddMissing -eq $true) {
        Write-Debug -Message "Exporting config to '$drrConfigFilePath'"
        $drrConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $drrConfigFilePath -Force
    }

    #Return true if all test pass.
    Write-Output -InputObject $true
}
