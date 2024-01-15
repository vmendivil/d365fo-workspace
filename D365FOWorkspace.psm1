# Depends on d365fo.tools

# Load defaults
$configFile = Get-ChildItem "$PSScriptRoot\config.ps1"
. $configFile.FullName

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

function Compare-FOWebConfigFile
{
    <#
    .SYNOPSIS
    Compares specific properties for the web.config file
    #>

    param (
        [string] $BackupDir
    )

    # Load DEV config file
    $devConfigPath = $env:USERPROFILE + '\Documents\Visual Studio Dynamics 365\DynamicsDevConfig.xml'

    if (-Not (Test-Path $devConfigPath))
    {
        Write-Host "Error: Dynamics DEV config file was not found in $env:USERPROFILE."
        return
    }

    [xml]$devConfig = Get-Content $devConfigPath
    $webRoot = $devConfig.DynamicsDevConfig.WebRoleDeploymentFolder
    $webConfigPath = Join-Path $webRoot 'web.config'

    # Compare web.config file
    $backupWebConfigPath = Join-Path $BackupDir ('web.config')

    # Step 1: Ensure both files exist
    if (!(Test-Path $webConfigPath -PathType Leaf) -or !(Test-Path $backupWebConfigPath -PathType Leaf))
    {
        Write-Host "Error: File $webConfigPath or Backup file $backupWebConfigPath does not exist."
        return
    }

    # Step 2: Compare specific properties (adjust as needed)
    $propertiesToCompare = @(
        "Aos.MetadataDirectory",
        "Aos.PackageDirectory",
        "bindir",
        "Common.BinDir",
        "Microsoft.Dynamics.AX.AosConfig.AzureConfig.bindir",
        "Common.DevToolsBinDir"
    )

    # Step 3: Load the content of both files as XML
    [xml]$xmlContent = Get-Content $webConfigPath
    [xml]$backupContent = Get-Content $backupWebConfigPath

    foreach ($property in $propertiesToCompare)
    {
        # Step 4: Check if the property exists in the original file
        $propertyNode = $xmlContent.configuration.appSettings.SelectSingleNode("add[@key='$property']")

        if ($propertyNode -ne $null) {
            # Step 5: Retrieve values for comparison from the original file
            $value = $propertyNode.Value

            # Step 6: Retrieve values for comparison from the backup file
            $backupValueNode = $backupContent.configuration.appSettings.SelectSingleNode("add[@key='$property']")
            
            if ($backupValueNode -ne $null) {
                $backupValue = $backupValueNode.Value

                # Step 7: Display the comparison result
                Write-Host "   Property: $property"
                Write-Host "      Current Value   : $value"
                Write-Host "      Backup Value    : $backupValue"
                Write-Host "      Values Match?   : $($value -eq $backupValue)"
                Write-Host ""
            } else {
                # Step 8: Display a message if the property does not exist in the backup file
                Write-Host "   Property: $property"
                Write-Host "      Does not exist in $backupWebConfigPath."
                Write-Host ""
            }
        } else {
            # Step 9: Display a message if the property does not exist in the original file
            Write-Host "   Property: $property"
            Write-Host "      Does not exist in $webConfigPath."
            Write-Host ""
        }
    }
}
