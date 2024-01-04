# Depends on d365fo.tools

# Load defaults
$configFile = Get-ChildItem "$PSScriptRoot\config.ps1"
. $configFile.FullName

# Define backup suffix
$backupSuffix = "_OrigBackup"

function Switch-FOWorkspace
{
	<#
	.SYNOPSIS
	Tells F&O (and F&O tools for Visual Studio) where to find source control and binaries.
	#>
	
	param (
		[string] $WorkspaceDir # = (Get-Location)
	)

	# You can modify these variables or expose them as function parameters
	$switchPackages = $true
	$switchVsDefaultProjectsPath = $true

	[string]$workspaceMetaDir = Join-Path $WorkspaceDir 'Metadata'
	[string]$workspaceProjectsDir = Join-Path $WorkspaceDir 'Projects'

	# Load DEV config file
	$devConfigPath = $env:USERPROFILE + '\Documents\Visual Studio Dynamics 365\DynamicsDevConfig.xml'

	if (-Not (Test-Path $devConfigPath))
	{
		throw 'Dynamics DEV config file was not found.'
	}

	[xml]$devConfig = Get-Content $devConfigPath
	$webRoot = $devConfig.DynamicsDevConfig.WebRoleDeploymentFolder
	$webConfigPath = $webRoot + '\web.config'
	[xml]$webConfig = Get-Content $webConfigPath

	$activeWorkspace = Get-FOWorkspace
	Write-Host "Previous workspace: $activeWorkspace"

	# Update path to packages
	if ($switchPackages)
	{
		Stop-D365Environment -ShowOriginalProgress
		
		$appSettings = $webConfig.configuration.appSettings
		$appSettings.SelectSingleNode("add[@key='Aos.MetadataDirectory']").Value = $workspaceMetaDir
		$appSettings.SelectSingleNode("add[@key='Aos.PackageDirectory']").Value = $workspaceMetaDir 
		$appSettings.SelectSingleNode("add[@key='bindir']").Value = $workspaceMetaDir
		$appSettings.SelectSingleNode("add[@key='Common.BinDir']").Value = $workspaceMetaDir
		$appSettings.SelectSingleNode("add[@key='Microsoft.Dynamics.AX.AosConfig.AzureConfig.bindir']").Value = $workspaceMetaDir
		$appSettings.SelectSingleNode("add[@key='Common.DevToolsBinDir']").Value = (Join-Path $workspaceMetaDir 'bin').ToString()

		$webConfig.Save($webConfigPath)
	}

	$activeWorkspace = Get-FOWorkspace
	Write-Host "Active workspace: $activeWorkspace"

	# Switch the default path for new projects in Visual Studio
	if ($switchVsDefaultProjectsPath)
	{
		$versionNum = ""

		switch ($VSVersion) {
			"2017" { $versionNum = "15" }
			"2019" { $versionNum = "16" }
			Default { $versionNum = "17" } # 2022
		}

		$settingsFilePattern = "$($env:LocalAppData)\Microsoft\VisualStudio\$versionNum*\Settings\CurrentSettings.vssettings"
		$settingsFile = Get-ChildItem $settingsFilePattern | Select-Object -First 1

		if ($settingsFile)
		{
			[xml]$vsConfigXml = Get-Content $settingsFile
			$options = $vsConfigXml.UserSettings.ToolsOptions
			$envOptions = $options.SelectSingleNode("ToolsOptionsCategory[@name='Environment']")
			$projOptions = $envOptions.SelectSingleNode("ToolsOptionsSubCategory[@name='ProjectsAndSolution']")
			$projPathElement = $projOptions.SelectSingleNode("PropertyValue[@name='ProjectsLocation']")

			if ($projPathElement.InnerText -ne $workspaceProjectsDir)
			{
				$projPathElement.InnerText = $workspaceProjectsDir
				$vsConfigXml.Save($settingsFile)

				Write-Information "Visual Studio configuration file updated"
			}
			else 
			{
				Write-Warning "Visual Studio configuration file project location not updated"
			}
		}
		else
		{
			Write-Warning "Visual Studio configuration file was not found."
		}
	}
}

function Get-FOWorkspace
{
	<#
	.SYNOPSIS
	Shows the model store folder currently used by F&O.
	#>
	
	(Get-D365EnvironmentSettings).Aos.MetadataDirectory
}

function Add-FOPackageSymLinks
{
	<#
	.SYNOPSIS
	Creates symbolic links for standard packages in D365FO.
	.DESCRIPTION
	Creates symbolic links for standard packages in D365FO. It requires admin permissions.
	Because only custom packages are stored in version control, downloading code from Azure DevOps
	creates a Metadata folder without standard packages (e.g. ApplicationSuite).
	To be able to connect F&O (and F&O DEV tools) to this folder, we need standard packages as well.
	One option would be copying standard packages there, but it's slow, it would waste disk space
	and it would cause problems with updates.
	Instead, we just create symbolic links to the folders.
	#>
	
	param (
		[string] $WorkspaceDir # = (Get-Location)
	)

	if ((Test-Path $WorkspaceDir) -ne $true)
	{
		throw "Path does not exist."
	}

	if ((Split-Path $WorkspaceDir  -Leaf) -eq 'Metadata')
	{
		$workspaceMetaDir = $WorkspaceDir;

	}
	else
	{
		$workspaceMetaDir = Join-Path $WorkspaceDir 'Metadata'
	}

	# For each folder in packages (it assumes that there are no custom packages)
	foreach ($dir in (Get-ChildItem (Get-FOPackagesDir) -dir))
	{	
		$targetDir = (Join-Path $workspaceMetaDir $dir.Name)

		if (Test-Path $targetDir)
		{
			#If there is a already a folder of the same name as the symblic link, delete it
			cmd /c rmdir /s /q $targetDir
		}

		# Create a symbolic link
		New-Item -ItemType SymbolicLink -Path $targetDir -Target $dir.FullName -Force
	}
}

function Get-FOPackagesDir
{
	<#
	.SYNOPSIS
	Tries to find the location of F&O model store (PackagesLocalDirectory).
	#>
	
	if ($PackageDir -and (Test-Path -Path $PackageDir))
	{
		return $PackageDir
	}

	foreach ($drive in (Get-Volume | where OperationalStatus -eq OK | where DriveLetter -ne $null | select -Expand DriveLetter))
	{
		$path = "${drive}:\AosService\PackagesLocalDirectory"
		
		if (Test-Path -Path $path)
		{
			return $path
		}
	}

	throw "Cannot find PackagesLocalDirectory. Specify the path in $($configFile.FullName)."
}

function Backup-FOConfigFiles
{
    <#
    .SYNOPSIS
    Backup all original files
    #>

    # Load D365FODEV config file
    $devConfigPath = $env:USERPROFILE + '\Documents\Visual Studio Dynamics 365\DynamicsDevConfig.xml'

    if (-Not (Test-Path $devConfigPath))
    {
        throw 'Dynamics DEV config file was not found.'
    }

    Backup-FOFile -filePath $devConfigPath

    # Load WEBROOT config file
    [xml]$devConfig = Get-Content $devConfigPath
    $webRoot = $devConfig.DynamicsDevConfig.WebRoleDeploymentFolder
    $webConfigPath = $webRoot + '\web.config'

    Backup-FOFile -filePath $webConfigPath

    # Load VS config file
    $versionNum = ""

    switch ($VSVersion) {
        "2017" { $versionNum = "15" }
        "2019" { $versionNum = "16" }
        Default { $versionNum = "17" } # 2022
    }

    $settingsFilePattern = "$($env:LocalAppData)\Microsoft\VisualStudio\$versionNum*\Settings\CurrentSettings.vssettings"
    $settingsFile = Get-ChildItem $settingsFilePattern | Select-Object -First 1

    Backup-FOFile -filePath $settingsFile
}

function Backup-FOFile
{
    <#
    .SYNOPSIS
    Check if a file backup exists, otherwise creates one
    #>

    param(
        [string] $filePath
    )

    # Step 1: Ensure the file exists
    if (-not (Test-Path $filePath -PathType Leaf)) {
        Write-Host "Error: The file does not exist at $filePath."
        return
    }

    # Step 2: Check if a backup already exists
    $directory = (Get-Item $filePath).Directory.FullName
    $fileName = (Get-Item $filePath).BaseName
    $extension = (Get-Item $filePath).Extension

    $backupFileName = "$fileName$backupSuffix$extension"
    $backupFilePath = Join-Path -Path $directory -ChildPath $backupFileName

    if (Test-Path $backupFilePath -PathType Leaf) {
        Write-Host "Backup already exists at $backupFilePath. Skipping backup creation."
    }
    else {
        # Step 3: Create the backup file
        Copy-Item $filePath $backupFilePath
        Write-Host "Backup created at $backupFilePath."
    }
}

function Restore-FOFile
{
    <#
    .SYNOPSIS
    Restore a single file from backup
    #>

    param(
        [string] $filePath
    )

    # Step 1: Check if a backup exists
    $directory = (Get-Item $filePath).Directory.FullName
    $fileName = (Get-Item $filePath).BaseName
    $extension = (Get-Item $filePath).Extension

    $backupFileName = "$fileName$backupSuffix$extension"
    $backupFilePath = Join-Path -Path $directory -ChildPath $backupFileName

    if (Test-Path $backupFilePath -PathType Leaf) {
        # Step 2: Restore the file from backup
        Copy-Item $backupFilePath $filePath -Force
        Write-Host "File restored from backup. Original file: $filePath, Backup file: $backupFilePath."
    }
    else {
        Write-Host "Error: Backup file does not exist for $filePath. Unable to restore."
    }
}

function Restore-FOConfigFiles
{
    <#
    .SYNOPSIS
    Restore all original files
    #>

    param(
        [switch]$IncludeVisualStudioFiles
    )

    # Load D365FODEV config file
    $devConfigPath = $env:USERPROFILE + '\Documents\Visual Studio Dynamics 365\DynamicsDevConfig.xml'

    if (-Not (Test-Path $devConfigPath))
    {
        throw 'Dynamics DEV config file was not found.'
    }

    Restore-FOFile -filePath $devConfigPath

    # Load WEBROOT config file
    [xml]$devConfig = Get-Content $devConfigPath
    $webRoot = $devConfig.DynamicsDevConfig.WebRoleDeploymentFolder
    $webConfigPath = $webRoot + '\web.config'

    Restore-FOFile -filePath $webConfigPath

    # Load VS config file if IncludeVisualStudioFiles is set
    if ($IncludeVisualStudioFiles) {
        $versionNum = ""

        switch ($VSVersion) {
            "2017" { $versionNum = "15" }
            "2019" { $versionNum = "16" }
            Default { $versionNum = "17" } # 2022
        }

        $settingsFilePattern = "$($env:LocalAppData)\Microsoft\VisualStudio\$versionNum*\Settings\CurrentSettings.vssettings"
        $settingsFile = Get-ChildItem $settingsFilePattern | Select-Object -First 1

        Restore-FOFile -filePath $settingsFile
    }
    else {
        Write-Host "Visual Studio files are excluded from restoration."
    }
}

function Remove-FOBackupFile
{
    <#
    .SYNOPSIS
    Delete a single backup file
    #>

    param(
        [string] $filePath
    )

    # Step 1: Check if a backup exists
    $directory = (Get-Item $filePath).Directory.FullName
    $fileName = (Get-Item $filePath).BaseName
    $extension = (Get-Item $filePath).Extension

    $backupFileName = "$fileName$backupSuffix$extension"
    $backupFilePath = Join-Path -Path $directory -ChildPath $backupFileName

    if (Test-Path $backupFilePath -PathType Leaf) {
        # Step 2: Delete the backup file
        Remove-Item $backupFilePath -Force
        Write-Host "Backup file deleted: $backupFilePath."
    }
    else {
        Write-Host "Error: Backup file does not exist for $filePath. Unable to delete."
    }
}

function Remove-AllFOBackupFiles
{
    <#
    .SYNOPSIS
    Delete all backup files with confirmation prompt
    #>

    param(
        [switch]$IncludeVisualStudioFiles
    )

    # Confirm deletion of all backup files
    $confirmation = Read-Host "Are you sure you want to delete all backup files? (Y/N)"
    if ($confirmation -ne 'Y') {
        Write-Host "Deletion canceled."
        return
    }

    # Define paths used in Backup and Restore functions
    $devConfigPath = $env:USERPROFILE + '\Documents\Visual Studio Dynamics 365\DynamicsDevConfig.xml'
    [xml]$devConfig = Get-Content $devConfigPath
    $webRoot = $devConfig.DynamicsDevConfig.WebRoleDeploymentFolder
    $webConfigPath = $webRoot + '\web.config'

    # Delete all backup files
    $backupFiles = @(
        Join-Path -Path (Get-Item $devConfigPath).Directory.FullName -ChildPath ($devConfig.Name + $backupSuffix),
        Join-Path -Path (Get-Item $webConfigPath).Directory.FullName -ChildPath ('web' + $backupSuffix + '.config')
    )

    if ($IncludeVisualStudioFiles) {
        $versionNum = ""
        switch ($VSVersion) {
            "2017" { $versionNum = "15" }
            "2019" { $versionNum = "16" }
            Default { $versionNum = "17" } # 2022
        }
        $settingsFilePattern = "$($env:LocalAppData)\Microsoft\VisualStudio\$versionNum*\Settings\CurrentSettings.vssettings"
        $visualStudioBackupFile = (Get-ChildItem $settingsFilePattern | Select-Object -First 1).FullName + $backupSuffix

        $backupFiles += $visualStudioBackupFile
    }

    foreach ($file in $backupFiles) {
        if (Test-Path $file -PathType Leaf) {
            Remove-Item $file -Force
            Write-Host "Backup file deleted: $file"
        }
        else {
            Write-Host "Error: Backup file does not exist at $file. Unable to delete."
        }
    }
}
