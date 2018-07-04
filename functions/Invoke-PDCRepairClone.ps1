function Invoke-PDCRepairClone {
<#
.SYNOPSIS
    Invoke-PDCRepairClone repairs the clones

.DESCRIPTION
    Invoke-PDCRepairClone has the ability to repair the clones when they have gotten disconnected from the image.
    In such a case the clone is no longer available for the database server and the database will either not show
    any information or the database will have the status (Recovery Pending).

    By running this command all the clones will be retrieved from the database for a certain host.

.PARAMETER HostName
    Set on or more hostnames to retrieve the configurations for

.PARAMETER SqlCredential
    Allows you to login to servers using SQL Logins as opposed to Windows Auth/Integrated/Trusted. To use:

    $scred = Get-Credential, then pass $scred object to the -SqlCredential parameter.

    Windows Authentication will be used if SqlCredential is not specified. SQL Server does not accept Windows credentials being passed as credentials.
    To connect as a different Windows user, run PowerShell as that user.

.PARAMETER EnableException
    By default, when something goes wrong we try to catch it, interpret it and give you a friendly warning message.
    This avoids overwhelming you with "sea of red" exceptions, but is inconvenient because it basically disables advanced scripting.
    Using this switch turns this "nice by default" feature off and enables you to catch exceptions with your own try/catch.

.PARAMETER WhatIf
    If this switch is enabled, no actions are performed but informational messages will be displayed that explain what would happen if the command were to run.

.PARAMETER Confirm
    If this switch is enabled, you will be prompted for confirmation before executing any operations that change state.

.NOTES
    Author: Sander Stad (@sqlstad, sqlstad.nl)

    Website: https://psdatabaseclone.io
    Copyright: (C) Sander Stad, sander@sqlstad.nl
    License: MIT https://opensource.org/licenses/MIT

.LINK
    https://psdatabaseclone.io/

.EXAMPLE
    Invoke-PDCRepairClone -Hostname Host1

    Repair the clones for Host1

#>
    [CmdLetBinding()]

    param(
        [Parameter(Mandatory = $true)]
        [string[]]$HostName,
        [System.Management.Automation.PSCredential]
        $SqlCredential,
        [switch]$EnableException
    )

    begin {
        # Test the module database setup
        try {
            Test-PDCConfiguration -EnableException
        }
        catch {
            Stop-PSFFunction -Message "Something is wrong in the module configuration" -ErrorRecord $_ -Continue
        }

        $pdcSqlInstance = Get-PSFConfigValue -FullName psdatabaseclone.database.server
        $pdcDatabase = Get-PSFConfigValue -FullName psdatabaseclone.database.name

    }

    process {
        # Test if there are any errors
        if (Test-PSFFunctionInterrupt) { return }

        # Loop through each of the hosts
        foreach ($hst in $HostName) {

            $query = "
                SELECT i.ImageLocation,
                        c.CloneLocation,
                        c.SqlInstance,
                        c.DatabaseName,
                        c.IsEnabled
                FROM dbo.Clone AS c
                    INNER JOIN dbo.Image AS i
                        ON i.ImageID = c.ImageID
                    INNER JOIN dbo.Host AS h
                        ON h.HostID = c.HostID
                WHERE h.HostName = '$hst';
            "

            Write-PSFMessage -Message "Query Host Clones`n$query" -Level Debug

            # Get the clones registered for the host
            try {
                Write-PSFMessage -Message "Get the clones for host $hst" -Level Verbose
                $results = Invoke-DbaSqlQuery -SqlInstance $pdcSqlInstance -Database $pdcDatabase -Query $query
            }
            catch {
                Stop-PSFFunction -Message "Couldn't get the clones for $hst" -Target $pdcSqlInstance -ErrorRecord $_ -Continue
            }

            # Loop through the results
            foreach ($result in $results) {

                # Get the databases
                Write-PSFMessage -Message "Retrieve the databases for $($result.SqlInstance)" -Level Verbose
                $databases = Get-DbaDatabase -SqlInstance $result.SqlInstance -SqlCredential $SqlCredential

                # Check if the parent of the clone can be reached
                if (Test-Path -Path $result.ImageLocation) {

                    # Mount the clone
                    try {
                        Write-PSFMessage -Message "Mounting vhd $($result.CloneLocation)" -Level Verbose

                        Mount-VHD -Path $result.CloneLocation -NoDriveLetter -ErrorAction SilentlyContinue
                    }
                    catch {
                        Stop-PSFFunction -Message "Couldn't mount vhd" -Target $clone -Continue
                    }
                }
                else {
                    Stop-PSFFunction -Message "Vhd $($result.CloneLocation) cannot be mounted because parent path cannot be reached" -Target $clone -Continue
                }

                # Check if the database is already attached
                if ($result.DatabaseName -notin $databases.Name) {

                    # Get all the files of the database
                    $databaseFiles = Get-ChildItem -Path $result.AccessPath -Recurse | Where-Object {-not $_.PSIsContainer}

                    # Setup the database filestructure
                    $dbFileStructure = New-Object System.Collections.Specialized.StringCollection

                    # Loop through each of the database files and add them to the file structure
                    foreach ($dbFile in $databaseFiles) {
                        $dbFileStructure.Add($dbFile.FullName) | Out-Null
                    }

                    Write-PSFMessage -Message "Mounting database from clone" -Level Verbose

                    # Mount the database using the config file
                    $null = Mount-DbaDatabase -SqlInstance $result.SQLInstance -Database $result.DatabaseName -FileStructure $dbFileStructure
                }
                else {
                    Write-PSFMessage -Message "Database $($result.Database) is already attached" -Level Verbose
                }

            }

        }

    }

    end {
        # Test if there are any errors
        if (Test-PSFFunctionInterrupt) { return }

        Write-PSFMessage -Message "Finished repairing clones" -Level Verbose
    }

}