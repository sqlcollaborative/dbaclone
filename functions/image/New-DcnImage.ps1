function New-DcnImage {
    <#
    .SYNOPSIS
        New-DcnImage creates a new image

    .DESCRIPTION
        New-DcnImage will create a new image based on a SQL Server database

        The command will either create a full backup or use the last full backup to create the image.

        Every image is created with the name of the database and a time stamp yyyyMMddHHmmss i.e "DB1_20180622171819.vhdx"

    .PARAMETER SourceSqlInstance
        Source SQL Server name or SMO object representing the SQL Server to connect to.
        This will be where the database is currently located

    .PARAMETER SourceSqlCredential
        Allows you to login to servers using SQL Logins as opposed to Windows Auth/Integrated/Trusted. To use:

        $scred = Get-Credential, then pass $scred object to the -SourceSqlCredential parameter.

        Windows Authentication will be used if SqlCredential is not specified. SQL Server does not accept Windows credentials being passed as credentials.
        To connect as a different Windows user, run PowerShell as that user.

    .PARAMETER SourceCredential
        Allows you to login to other parts of a system like folders. To use:

        $scred = Get-Credential, then pass $scred object to the -SourceCredential parameter.

    .PARAMETER DestinationSqlInstance
        SQL Server name or SMO object representing the SQL Server to connect to.
        This is the server to use to temporarily restore the database to create the image.

    .PARAMETER DestinationSqlCredential
        Allows you to login to servers using SQL Logins as opposed to Windows Auth/Integrated/Trusted. To use:

        $scred = Get-Credential, then pass $scred object to the -DestinationSqlCredential parameter.

        Windows Authentication will be used if SqlCredential is not specified. SQL Server does not accept Windows credentials being passed as credentials.
        To connect as a different Windows user, run PowerShell as that user.

    .PARAMETER DestinationCredential
        Allows you to login to other parts of a system like folders. To use:

        $scred = Get-Credential, then pass $scred object to the -DestinationCredential parameter.

    .PARAMETER DcnSqlCredential
        Allows you to login to servers using SQL Logins as opposed to Windows Auth/Integrated/Trusted.
        This works similar as SqlCredential but is only meant for authentication to the dbaclone database server and database.

        By default the script will try to retrieve the configuration value "dbaclone.informationstore.credential"

    .PARAMETER ImageNetworkPath
        Network path where to save the image. This has to be a UNC path

    .PARAMETER ImageLocalPath
        Local path where to save the image

    .PARAMETER Database
        Databases to create an image of

    .PARAMETER VhdType
        The type of the harddisk. This can either by VHD (version 1) or VHDX (version 2)
        The default is VHDX.

    .PARAMETER CreateFullBackup
        Create a new full backup of the database. The backup will be saved in the default backup directory

    .PARAMETER UseLastFullBackup
        Use the last full backup created for the database

    .PARAMETER BackupFilePath
        Use a specific backup file to create the image

    .PARAMETER CopyOnlyBackup
        Create a backup as COPY_ONLY

    .PARAMETER Force
        Forcefully execute commands when needed

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

        Website: https://psdatabaseclone.org
        Copyright: (C) Sander Stad, sander@sqlstad.nl
        License: MIT https://opensource.org/licenses/MIT

    .LINK
        https://psdatabaseclone.org/

    .EXAMPLE
        New-DcnImage -SourceSqlInstance SQLDB1 -DestinationSqlInstance SQLDB2 -ImageLocalPath C:\Temp\images\ -Database DB1 -CreateFullBackup

        Create an image for databas DB1 from SQL Server SQLDB1. The temporary destination will be SQLDB2.
        The image will be saved in C:\Temp\images.
    .EXAMPLE
        New-DcnImage -SourceSqlInstance SQLDB1 -DestinationSqlInstance SQLDB2 -ImageLocalPath C:\Temp\images\ -Database DB1 -UseLastFullBackup

        Create an image from the database DB1 on SQLDB1 using the last full backup and use SQLDB2 as the temporary database server.
        The image is written to c:\Temp\images
    #>
    [CmdLetBinding(SupportsShouldProcess = $true)]

    param(
        # [parameter(Mandatory = $true)]
        # [ValidateNotNullOrEmpty()]
        [DbaInstanceParameter]$SourceSqlInstance,
        [PSCredential]$SourceSqlCredential,
        [PSCredential]$SourceCredential,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [DbaInstanceParameter]$DestinationSqlInstance,
        [PSCredential]$DestinationSqlCredential,
        [PSCredential]$DestinationCredential,
        [PSCredential]$DcnSqlCredential,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object[]]$Database,
        [string]$ImageNetworkPath,
        [string]$ImageLocalPath,
        [ValidateSet('VHD', 'VHDX', 'vhd', 'vhdx')]
        [string]$VhdType,
        [switch]$CreateFullBackup,
        [switch]$UseLastFullBackup,
        [string]$BackupFilePath,
        [switch]$CopyOnlyBackup,
        [Alias('MaskingConfigFile', 'MaskingConfigFilePath')]
        [switch]$Force,
        [switch]$EnableException
    )

    begin {
        # Check if the console is run in Administrator mode
        if ( -not (Test-PSFPowerShell -Elevated) ) {
            Stop-PSFFunction -Message "Module requires elevation. Please run the console in Administrator mode" -Continue
        }

        # Check if the setup has ran
        if (-not (Test-DcnModule -SetupStatus)) {
            Stop-PSFFunction -Message "The module setup has NOT yet successfully run. Please run 'Set-DcnConfiguration'" -Continue
        }

        if(-not $CreateFullBackup -and -not $UseLastFullBackup -and -not $BackupFilePath){
            Stop-PSFFunction -Message "Unable to get last backup file. Please use -CreateFullBackup, -UseLastFullBackup or -BackupFile" -Continue
        }

        # Checking parameters
        if (-not $ImageNetworkPath) {
            Stop-PSFFunction -Message "Please enter the network path where to save the images" -Continue
        }

        # Check the vhd type
        if (-not $VhdType) {
            Write-PSFMessage -Message "Setting vhd type to 'VHDX'" -Level Verbose
            $VhdType = 'VHDX'
        }

        # Get the information store
        $informationStore = Get-PSFConfigValue -FullName dbaclone.informationstore.mode

        if ($informationStore -eq 'SQL') {
            # Get the module configurations
            $pdcSqlInstance = Get-PSFConfigValue -FullName dbaclone.database.Server
            $pdcDatabase = Get-PSFConfigValue -FullName dbaclone.database.name
            if (-not $DcnSqlCredential) {
                $pdcCredential = Get-PSFConfigValue -FullName dbaclone.informationstore.credential -Fallback $null
            }
            else {
                $pdcCredential = $DcnSqlCredential
            }

            # Test the module database setup
            if ($PSCmdlet.ShouldProcess("Test-DcnConfiguration", "Testing module setup")) {
                try {
                    Test-DcnConfiguration -SqlCredential $pdcCredential -EnableException
                }
                catch {
                    Stop-PSFFunction -Message "Something is wrong in the module configuration" -ErrorRecord $_ -Continue
                }
            }
        }

        Write-PSFMessage -Message "Started image creation" -Level Verbose

        # Try connecting to the instance
        if (-Not($BackupFilePath)) 
        {
            Write-PSFMessage -Message "Attempting to connect to Sql Server $SourceSqlInstance.." -Level Verbose
            try {
                $sourceServer = Connect-DbaInstance -SqlInstance $SourceSqlInstance -SqlCredential $SourceSqlCredential
            }
            catch {
                Stop-PSFFunction -Message "Could not connect to Sql Server instance $SourceSqlInstance" -ErrorRecord $_ -Target $SourceSqlInstance
                return
            }
        }

        # Cleanup the values in the network path
        if ($ImageNetworkPath.EndsWith("\")) {
            $ImageNetworkPath = $ImageNetworkPath.Substring(0, $ImageNetworkPath.Length - 1)
        }

        # Make up the data from the network path
        try {
            [uri]$uri = New-Object System.Uri($ImageNetworkPath)
            $uriHost = $uri.Host
        }
        catch {
            Stop-PSFFunction -Message "The image network path $ImageNetworkPath is not valid" -ErrorRecord $_ -Target $ImageNetworkPath
            return
        }

        # Setup the computer object
        $computer = [PsfComputer]$uriHost

        if (-not $computer.IsLocalhost) {
            # Get the result for the remote test
            $resultPSRemote = Test-DcnRemoting -ComputerName $computer -Credential $Credential

            # Check the result
            if ($resultPSRemote.Result) {
                $command = [scriptblock]::Create("Import-Module dbaclone -Force")

                try {
                    Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $Credential
                }
                catch {
                    Stop-PSFFunction -Message "Couldn't import module remotely" -Target $command
                    return
                }
            }
            else {
                Stop-PSFFunction -Message "Couldn't connect to host remotely.`nVerify that the specified computer name is valid, that the computer is accessible over the network, and that a firewall exception for the WinRM service is enabled and allows access from this computer" -Target $resultPSRemote -Continue
            }
        }

        # Get the local path from the network path
        if (-not $ImageLocalPath) {
            if ($PSCmdlet.ShouldProcess($ImageNetworkPath, "Converting UNC path to local path")) {
                try {
                    # Check if computer is local
                    if ($computer.IsLocalhost) {
                        $ImageLocalPath = Convert-DcnUncPathToLocalPath -UncPath $ImageNetworkPath -EnableException
                    }
                    else {
                        $command = "Convert-DcnUncPathToLocalPath -UncPath `"$ImageNetworkPath`" -EnableException"
                        $commandGetLocalPath = [ScriptBlock]::Create($command)
                        $ImageLocalPath = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $commandGetLocalPath -Credential $DestinationCredential

                        if (-not $ImageLocalPath) {
                            Stop-PSFFunction -Message "Could not convert network path to local path" -Target $ImageLocalPath
                            return
                        }
                    }

                    Write-PSFMessage -Message "Converted '$ImageNetworkPath' to '$ImageLocalPath'" -Level Verbose

                }
                catch {
                    Stop-PSFFunction -Message "Something went wrong getting the local image path" -Target $ImageNetworkPath
                    return
                }
            }
        }
        else {
            # Cleanup the values in the network path
            if ($ImageLocalPath.EndsWith("\")) {
                $ImageLocalPath = $ImageLocalPath.Substring(0, $ImageLocalPath.Length - 1)
            }

            # Check if the assigned value in the local path corresponds to the one retrieved
            try {
                # Check if computer is local
                if ($computer.IsLocalhost) {
                    $convertedLocalPath = Convert-DcnUncPathToLocalPath -UncPath $ImageNetworkPath -EnableException
                }
                else {
                    $command = "Convert-DcnUncPathToLocalPath -UncPath `"$ImageNetworkPath`" -EnableException"
                    $commandGetLocalPath = [ScriptBlock]::Create($command)
                    $convertedLocalPath = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $commandGetLocalPath -Credential $DestinationCredential
                }

                Write-PSFMessage -Message "Converted '$ImageNetworkPath' to '$ImageLocalPath'" -Level Verbose

                # Check if the ImageLocalPath and convertedLocalPath are the same
                if ($ImageLocalPath -ne $convertedLocalPath) {
                    Stop-PSFFunction -Message "The local path '$ImageLocalPath' is not the same location as the network path '$ImageNetworkPath'" -Target $ImageNetworkPath
                    return
                }

            }
            catch {
                Stop-PSFFunction -Message "Something went wrong getting the local image path" -Target $ImageNetworkPath
                return
            }
        }

        if (-Not($BackupFilePath)) 
        {
            # Check the image local path
            if ($PSCmdlet.ShouldProcess("Verifying image local path")) {
                if ((Test-DbaPath -Path $ImageLocalPath -SqlInstance $DestinationSqlCredential -SqlCredential $DestinationCredential) -ne $true) {
                    Stop-PSFFunction -Message "Image local path $ImageLocalPath is not valid directory or can't be reached." -Target $DestinationSqlCredential
                    return
                }
            }
        }

        # Clean up the paths
        if ($ImageLocalPath.EndsWith("\")) {
            $ImageLocalPath = $ImageLocalPath.Substring(0, $ImageLocalPath.Length - 1)
        }

        $imagePath = $ImageLocalPath

        # Check the database parameter
        if (-Not($BackupFilePath)) 
        {
            if ($Database) {
                foreach ($db in $Database) {
                    if ($db -notin $sourceServer.Databases.Name) {
                        Stop-PSFFunction -Message "Database $db cannot be found on instance $SourceSqlInstance" -Target $SourceSqlInstance
                    }
                }

                $DatabaseCollection = $sourceServer.Databases | Where-Object { $_.Name -in $Database }
            }
            else {
                Stop-PSFFunction -Message "Please supply a database to create an image for" -Target $SourceSqlInstance -Continue
            }
        }
        else 
        {
            $DatabaseCollection = 
            @{
                Name = $Database
                Size = 1
            };
        }
            
        if($BackupFilePath){
            if(-not (Test-Path -Path $BackupFilePath)){
                Stop-PSFFunction -Message "Could not find backup file '$($BackupFilePath)'"
            }

            if($Database.Count -gt 1){
                Stop-PSFFunction -Message "You cannot enter multiple databases for the same backup file. Please just enter one"
            }
        }



        # Set time stamp
        $timestamp = Get-Date -format "yyyyMMddHHmmss"

    }

    process {
        # Test if there are any errors
        if (Test-PSFFunctionInterrupt) { return }

        # Loop through each of the databases
        foreach ($db in $DatabaseCollection) {
            if (-Not($BackupFilePath)) 
            {
                Write-PSFMessage -Message "Creating image for database $db from $SourceSqlInstance" -Level Verbose
            }
            else 
            {
                Write-PSFMessage -Message "Creating image for database $db from $BackupFilePath" -Level Verbose
            }

            if ($PSCmdlet.ShouldProcess($db, "Checking available disk space for database")) {
                # Check the database size to the available disk space
                if ($computer.IsLocalhost) {
                    $availableMB = (Get-PSDrive -Name $ImageLocalPath.Substring(0, 1)).Free / 1MB
                }
                else {
                    $command = [ScriptBlock]::Create("(Get-PSDrive -Name $($ImageLocalPath.Substring(0, 1)) ).Free / 1MB")
                    $availableMB = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $commandGetLocalPath -Credential $DestinationCredential
                }

                $dbSizeMB = $db.Size

                if ($availableMB -lt $dbSizeMB) {
                    Stop-PSFFunction -Message "Size of database $($db.Name) does not fit within the image local path" -Target $db -Continue
                }
            }

            # Setup the image variables
            $imageName = "$($db.Name)_$timestamp"

            # Setup the access path
            $accessPath = Join-PSFPath -Path $ImageLocalPath -Child $imageName

            # Setup the vhd path
            $vhdPath = "$($accessPath).$($VhdType.ToLower())"

            if ($CreateFullBackup) {
                if ($PSCmdlet.ShouldProcess($db, "Creating full backup for database $db")) {

                    # Create the backup
                    Write-PSFMessage -Message "Creating new full backup for database $db" -Level Verbose
                    $lastFullBackup = Backup-DbaDatabase -SqlInstance $SourceSqlInstance -SqlCredential $SourceSqlCredential -Database $db.Name -CopyOnly:$CopyOnlyBackup
                }
            }
            elseif ($BackupFilePath) {
                [pscustomobject]$lastFullBackup = @{
                    Path = $BackupFilePath
                }
            }
            else {
                Write-PSFMessage -Message "Trying to retrieve the last full backup for $db" -Level Verbose

                # Get the last full backup
                $lastFullBackup = Get-DbaDbBackupHistory -SqlInstance $SourceSqlInstance -SqlCredential $SourceSqlCredential -Database $db.Name -LastFull
            }

            if (-not $lastFullBackup.Path) {
                Stop-PSFFunction -Message "No full backup could be found. Please use -CreateFullBackup or create a full backup manually" -Target $lastFullBackup
                return
            }
            elseif (-not (Test-Path -Path $lastFullBackup.Path)) {
                Stop-PSFFunction -Message "Could not access the full backup file. Check if it exists or that you have enough privileges to access it" -Target $lastFullBackup
                return
            }

            if ($PSCmdlet.ShouldProcess("$imageName", "Creating the vhd")) {
                # try to create the new VHD
                try {
                    Write-PSFMessage -Message "Create the vhd $imageName" -Level Verbose

                    # Check if computer is local
                    if ($computer.IsLocalhost) {
                        $null = New-DcnVhdDisk -Destination $imagePath -Name $imageName -VhdType $VhdType -EnableException
                    }
                    else {
                        $command = [ScriptBlock]::Create("New-DcnVhdDisk -Destination '$imagePath' -Name $imageName -VhdType $VhdType -EnableException")
                        $null = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $DestinationCredential
                    }

                }
                catch {
                    Stop-PSFFunction -Message "Couldn't create vhd(x) $imageName" -Target $imageName -ErrorRecord $_ -Continue
                }
            }


            if ($PSCmdlet.ShouldProcess("$imageName", "Initializing the vhd")) {
                # Try to initialize the vhd
                try {
                    Write-PSFMessage -Message "Initializing the vhd $imageName" -Level Verbose

                    # Check if computer is local
                    if ($computer.IsLocalhost) {
                        $diskResult = Initialize-DcnVhdDisk -Path $vhdPath -Credential $DestinationCredential -EnableException
                    }
                    else {
                        $command = [ScriptBlock]::Create("Initialize-DcnVhdDisk -Path $vhdPath -EnableException")
                        $diskResult = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $DestinationCredential
                    }
                }
                catch {
                    Stop-PSFFunction -Message "Couldn't initialize vhd $vhdPath" -Target $imageName -ErrorRecord $_
                }
            }

            # Create folder structure for image
            $imageDataFolder = Join-PSFPath -Path $imagePath -Child "$($imageName)\Data"
            $imageLogFolder = Join-PSFPath -Path $imagePath -Child "$($imageName)\Log"

            # try to create access path
            try {
                # Check if access path is already present
                if (-not (Test-Path -Path $accessPath)) {
                    if ($PSCmdlet.ShouldProcess($accessPath, "Creating access path $accessPath")) {
                        try {
                            # Check if computer is local
                            if ($computer.IsLocalhost) {
                                $null = New-Item -Path $accessPath -ItemType Directory -Force
                            }
                            else {
                                $command = [ScriptBlock]::Create("New-Item -Path $accessPath -ItemType Directory -Force")
                                $null = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $DestinationCredential

                                # Set the permissions
                                $script = "
                                    `$permission = 'Everyone', 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'
                                    `$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule `$permission
                                    `$acl = Get-Acl -Path '$accessPath'
                                    `$acl.SetAccessRule(`$accessRule)
                                    `$acl | Set-Acl '$accessPath'
                                "

                                $command = [ScriptBlock]::Create($script)
                                $null = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $DestinationCredential
                            }
                        }
                        catch {
                            Stop-PSFFunction -Message "Couldn't create access path directory" -ErrorRecord $_ -Target $accessPath -Continue
                        }
                    }

                    # Set the permissions
                    #$permission = "Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
                    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "ContainerInherit,Objectinherit", "None", "Allow")
                    $acl = Get-Acl -Path $accessPath
                    $acl.SetAccessRule($accessRule)
                    Set-Acl -Path $accessPath -AclObject $acl
                }

                # Get the properties of the disk and partition
                $disk = $diskResult.Disk
                $partition = Get-Partition -DiskNumber $disk.Number | Where-Object { $_.Type -ne "Reserved" } | Select-Object -First 1

                if ($PSCmdlet.ShouldProcess($accessPath, "Adding access path '$accessPath' to mounted disk")) {
                    # Add the access path to the mounted disk
                    if ($computer.IsLocalhost) {
                        $null = Add-PartitionAccessPath -DiskNumber $disk.Number -PartitionNumber $partition.PartitionNumber -AccessPath $accessPath -ErrorAction SilentlyContinue
                    }
                    else {
                        $command = [ScriptBlock]::Create("Add-PartitionAccessPath -DiskNumber $($disk.Number) -PartitionNumber $($partition.PartitionNumber) -AccessPath $accessPath -ErrorAction SilentlyContinue")
                        $null = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $DestinationCredential
                    }
                }

            }
            catch {
                Stop-PSFFunction -Message "Couldn't create access path for partition" -ErrorRecord $_ -Target $diskResult.partition
            }

            # Check if image data folder exist
            if (-not (Test-Path -Path $imageDataFolder)) {
                if ($PSCmdlet.ShouldProcess($accessPath, "Creating data folder in vhd")) {
                    try {
                        Write-PSFMessage -Message "Creating data folder for image" -Level Verbose

                        # Check if computer is local
                        if ($computer.IsLocalhost) {
                            $null = New-Item -Path $imageDataFolder -ItemType Directory

                            $acl = Get-ACL -Path $imageDataFolder
                            $acl.SetAccessRule($accessRule)
                            Set-Acl -Path $imageDataFolder -AclObject $acl
                        }
                        else {
                            $command = [ScriptBlock]::Create("New-Item -Path $imageDataFolder -ItemType Directory")
                            $null = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $DestinationCredential
                        }
                    }
                    catch {
                        Stop-PSFFunction -Message "Couldn't create image data folder" -Target $imageName -ErrorRecord $_ -Continue
                    }
                }
            }

            # Test if the image log folder exists
            if (-not (Test-Path -Path $imageLogFolder)) {
                if ($PSCmdlet.ShouldProcess($accessPath, "Creating log folder in vhd")) {
                    try {
                        Write-PSFMessage -Message "Creating transaction log folder for image" -Level Verbose

                        # Check if computer is local
                        if ($computer.IsLocalhost) {
                            $null = New-Item -Path $imageLogFolder -ItemType Directory

                            $acl = Get-ACL -Path $imageLogFolder
                            $acl.SetAccessRule($accessRule)
                            Set-Acl -Path $imageLogFolder -AclObject $acl
                        }
                        else {
                            $command = [ScriptBlock]::Create("New-Item -Path $imageLogFolder -ItemType Directory")
                            $null = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $DestinationCredential
                        }

                    }
                    catch {
                        Stop-PSFFunction -Message "Couldn't create image data folder" -Target $imageName -ErrorRecord $_ -Continue
                    }
                }
            }

            # Setup the temporary database name
            $tempDbName = "$($db.Name)-dbaclone"

            if ($PSCmdlet.ShouldProcess($tempDbName, "Restoring database")) {
                # Restore database to image folder
                try {
                    Write-PSFMessage -Message "Restoring database $db on $DestinationSqlInstance" -Level Verbose

                    $params = @{
                        SqlInstance              = $DestinationSqlInstance
                        SqlCredential            = $DestinationSqlCredential
                        DatabaseName             = $tempDbName
                        Path                     = $lastFullBackup.Path
                        DestinationDataDirectory = $imageDataFolder
                        DestinationLogDirectory  = $imageLogFolder
                        WithReplace              = $true
                        EnableException          = $true
                    }

                    $restore = Restore-DbaDatabase @params
                }
                catch {
                    Stop-PSFFunction -Message "Couldn't restore database $db as $tempDbName on $DestinationSqlInstance.`n$($_)" -Target $restore -ErrorRecord $_ -Continue
                }
            }

            # Detach database
            if ($PSCmdlet.ShouldProcess($tempDbName, "Detaching database")) {
                try {
                    Write-PSFMessage -Message "Detaching database $tempDbName on $DestinationSqlInstance" -Level Verbose
                    $null = Dismount-DbaDatabase -SqlInstance $DestinationSqlInstance -Database $tempDbName -SqlCredential $DestinationSqlCredential -Force
                }
                catch {
                    Stop-PSFFunction -Message "Couldn't detach database $db as $tempDbName on $DestinationSqlInstance" -Target $db -ErrorRecord $_ -Continue
                }
            }

            if ($PSCmdlet.ShouldProcess($vhdPath, "Dismounting the vhd")) {
                # Dismount the vhd
                try {
                    Write-PSFMessage -Message "Dismounting vhd" -Level Verbose

                    # Check if computer is local
                    if ($computer.IsLocalhost) {
                        # Dismount the VHD
                        $null = Dismount-DiskImage -ImagePath $vhdPath

                        # Remove the access path
                        $null = Remove-Item -Path $accessPath -Recurse -Force
                    }
                    else {
                        $command = [ScriptBlock]::Create("Dismount-DiskImage -ImagePath $vhdPath")
                        $null = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $DestinationCredential

                        $command = [ScriptBlock]::Create("Remove-Item -Path $accessPath -Force")
                        $null = Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $DestinationCredential
                    }
                }
                catch {
                    Stop-PSFFunction -Message "Couldn't dismount vhd" -Target $imageName -ErrorRecord $_ -Continue
                }
            }

            # Write the data to the database
            $imageLocation = Join-PSFPath $uri.LocalPath -Child "$($imageName).vhdx"
            $sizeMB = $dbSizeMB
            $databaseName = $db.Name
            $databaseTS = $lastFullBackup.Start

            if ($informationStore -eq 'SQL') {
                $query = "
                DECLARE @ImageID INT;
                EXECUTE dbo.Image_New @ImageID = @ImageID OUTPUT,				  -- int
                                    @ImageName = '$($imageName)',                    -- varchar(100)
                                    @ImageLocation = '$($imageLocation)',			  -- varchar(255)
                                    @SizeMB = $($sizeMB),							  -- int
                                    @DatabaseName = '$($databaseName)',			  -- varchar(100)
                                    @DatabaseTimestamp = '$($databaseTS)'            -- datetime

                SELECT @ImageID as ImageID
            "

                # Add image to database
                if ($PSCmdlet.ShouldProcess($imageName, "Adding image to database")) {
                    try {
                        Write-PSFMessage -Message "Saving image information in database" -Level Verbose

                        $result += Invoke-DbaQuery -SqlInstance $pdcSqlInstance -SqlCredential $pdcCredential -Database $pdcDatabase -Query $query -EnableException
                        $imageID = $result.ImageID
                    }
                    catch {
                        Stop-PSFFunction -Message "Couldn't add image to database" -Target $imageName -ErrorRecord $_
                    }
                }
            }
            elseif ($informationStore -eq 'File') {
                [array]$images = $null

                # Get all the images
                try {
                    $images = Get-DcnImage
                }
                catch {
                    Stop-PSFFunction -Message "Couldn't get images" -Target $imageName -ErrorRecord $_
                    return
                }


                # Setup the new image id
                if ($images.Count -ge 1) {
                    $imageID = ($images[-1].ImageID | Sort-Object ImageID) + 1
                }
                else {
                    $imageID = 1
                }

                # Add the new information to the array
                $images += [PSCustomObject]@{
                    ImageID           = $imageID
                    ImageName         = $imageName
                    ImageLocation     = $imageLocation
                    SizeMB            = $sizeMB
                    DatabaseName      = $databaseName
                    DatabaseTimestamp = $databaseTS
                    CreatedOn         = (Get-Date -format "yyyyMMddHHmmss")
                }

                # Test if the JSON folder can be reached
                if (-not (Test-Path -Path "DCNJSONFolder:\")) {
                    $command = [scriptblock]::Create("Import-Module dbaclone -Force")

                    try {
                        Invoke-PSFCommand -ComputerName $computer -ScriptBlock $command -Credential $Credential
                    }
                    catch {
                        Stop-PSFFunction -Message "Couldn't import module remotely" -Target $command
                        return
                    }
                }

                # Set the image file
                $jsonImageFile = "DCNJSONFolder:\images.json"

                # Convert the data back to JSON
                $images | ConvertTo-Json | Set-Content $jsonImageFile
            }

            # Add the results to the custom object
            [PSCustomObject]@{
                ImageID           = $imageID
                ImageName         = $imageName
                ImageLocation     = $imageLocation
                SizeMB            = $sizeMB
                DatabaseName      = $databaseName
                DatabaseTimestamp = $databaseTS
            }
        } # for each database
    } # end process

    end {
        # Test if there are any errors
        if (Test-PSFFunctionInterrupt) { return }

        Write-PSFMessage -Message "Finished creating database image" -Level Verbose
    }

}
