<#
.SYNOPSIS
    Creates a new Win32 application in Microsoft Intune.

.DESCRIPTION
    This function packages an application installer as a .intunewin file and uploads it to Microsoft Intune as a Win32 app.
    It creates detection rules based on file existence and registry version, sets requirement rules for x64 architecture
    and Windows 10 20H2 or later, and configures the application with metadata like display name, publisher, description,
    icon, and installation behavior.

.PARAMETER SourcePath
    The file path to the main installer executable (.exe file). This file will be packaged into a .intunewin file.

.PARAMETER DisplayName
    The display name for the application as it will appear in the Company Portal.

.PARAMETER Publisher
    The publisher name of the application.

.PARAMETER Description
    The description of the application that will appear in the Company Portal.

.PARAMETER ApplicationExePath
    The full path to the main executable on the client machine. Used for file-based detection rule.

.PARAMETER RegKeyPath
    The registry key path that contains the DisplayVersion value on the client machine. Used for registry-based detection rule.

.PARAMETER DisplayVersion
    The application version value from the registry on the client machine. Used for version comparison in detection rules.

.PARAMETER IconPath
    The file path to the application icon image (.png or .jpg file).

.PARAMETER AppVersion
    The application version that will appear in the Company Portal. Defaults to the DisplayVersion parameter value.

.PARAMETER Developer
    The developer/author of the application that appears in the Company Portal. Defaults to the Publisher parameter value.

.PARAMETER Owner
    The product owner of the application within the organization. Defaults to an empty string.

.PARAMETER Notes
    Additional notes about the application. Defaults to an empty string.

.PARAMETER InformationURL
    A URL link to a knowledge base or information page about the application. Defaults to an empty string.

.PARAMETER PrivacyURL
    A URL link to the application's privacy policy. Defaults to an empty string.

.PARAMETER CompanyPortalFeaturedApp
    Specifies whether the application should be featured in the Company Portal. Defaults to $false.

.PARAMETER CategoryName
    One or more category names to organize the application in Intune. Optional parameter.

.PARAMETER RestartBehavior
    Defines the restart behavior after installation. Valid values are "allow", "basedOnReturnCode", "suppress", or "force". Defaults to "suppress".

.PARAMETER InstallExperience
    Specifies the installation context. Valid values are "SYSTEM" or "User". Defaults to "SYSTEM".

.PARAMETER InstallCommandLine
    The command line used to install the application. Defaults to "[SourceFileName] -DeploymentType Install".

.PARAMETER UninstallCommandLine
    The command line used to uninstall the application. Defaults to "[SourceFileName] -DeploymentType Uninstall".

.EXAMPLE
    New-IntuneApp -SourcePath "C:\Installers\MyApp\Deploy-Application.exe" `
                  -DisplayName "My Application" `
                  -Publisher "Contoso Ltd" `
                  -Description "My application for business users" `
                  -ApplicationExePath "C:\Program Files\MyApp\MyApp.exe" `
                  -RegKeyPath "HKEY_LOCAL_MACHINE\SOFTWARE\MyApp" `
                  -DisplayVersion "1.0.0" `
                  -IconPath "C:\Installers\MyApp\icon.png"

    Creates a new Intune Win32 app using the specified parameters with default installation behavior.

.EXAMPLE
    New-IntuneApp -SourcePath "C:\Installers\MyApp\Deploy-Application.exe" `
                  -DisplayName "My Application" `
                  -Publisher "Contoso Ltd" `
                  -Description "My application for business users" `
                  -ApplicationExePath "C:\Program Files\MyApp\MyApp.exe" `
                  -RegKeyPath "HKEY_LOCAL_MACHINE\SOFTWARE\MyApp" `
                  -DisplayVersion "2.0.0" `
                  -IconPath "C:\Installers\MyApp\icon.png" `
                  -CategoryName "Productivity", "Business Apps" `
                  -CompanyPortalFeaturedApp $true `
                  -RestartBehavior "allow"

    Creates a featured Intune Win32 app with multiple categories and allows restart after installation.

.NOTES
    Author: Matt McPhee
    Requires: IntuneWin32App PowerShell module
    Requires: Microsoft Graph authentication with appropriate permissions
    The function expects PSAppDeployToolkit-style command line parameters by default.
    The function creates detection rules based on both file existence and registry version.
    Created: 7-Jan-2026
    Updated: 8-Jan-2026

    Version History:
    1.0.0 - 8-Jan-2026 - Function created
    1.0.1 - 30-Jan-2026 - changed architecture requirement from x64 to AllWithARM64

.LINK
    https://github.com/MSEndpointMgr/IntuneWin32App
#>
function New-IntuneApp {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        # SourcePath - path to the main installer exe
        [Parameter(Mandatory)]
        [ValidateScript({
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "$_ could not be found or is not a path to a file."
            } else {
                return $true
            }
        })]
        [string]$SourcePath,

        # InstalledApplicationSizeMB
        [Parameter(Mandatory)]
        [int]$InstalledApplicationSizeMB,

        # DisplayName - desired displayname for the application in company portal
        [Parameter(Mandatory)]
        [string]$DisplayName,

        # Publisher - publisher of the application
        [Parameter(Mandatory)]
        [string]$Publisher,

        # Description - description that appears in company portal
        [Parameter(Mandatory)]
        [string]$Description,

        # ApplicationExePath - path to app launcher file on the client machine (will be used for detection)
        [Parameter(Mandatory = $false)]
        [string]$ApplicationLauncherPath,

        # RegKeyPath - registry path that holds desired value name on the client machine (will be used for detection)
        [Parameter(Mandatory)]
        [string]$RegKeyPath,

        # RegKeyValueName - name of the registry value to use for detection
        [Parameter(Mandatory)]
        [string]$RegKeyValueName,

        # RegKeyVersionValue - the version value that appears in the registry on the client machine (will be used for detection)
        [Parameter(Mandatory = $false)]
        [string]$RegKeyVersionValue,

        # IconPath - path to icon image file
        [Parameter(Mandatory=$false)]
        [ValidateScript({
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "$_ not found or is a folder instead of a file."
            } elseif ($_ -notlike "*.png" -and $_ -notlike "*.jpg") {
                throw "$_ is not a valid image file path ending in png or jpg."
            } else {
                return $true
            }
        })]
        [string]$IconPath,

        # Developer - author of the application (optional, defaults to publisher)
        [Parameter(Mandatory = $false)]
        [string]$Developer = $Publisher,

        # Owner - product owner of the application within the organization (optional)
        [Parameter(Mandatory = $false)]
        [string]$Owner,

        # Notes - info for fellow intune administrators (optional)
        [Parameter(Mandatory = $false)]
        [string]$Notes,

        # InformationURL - link to a knowledge base item with further info (optional)
        [Parameter(Mandatory = $false)]
        [string]$InformationURL,

        # PrivacyURL - link to the app's privacy policy (optional)
        [Parameter(Mandatory = $false)]
        [string]$PrivacyURL,

        # CompanyPortalFeaturedApp - whether to feature the app in company portal (optional, defaults to false)
        [Parameter(Mandatory = $false)]
        [bool]$CompanyPortalFeaturedApp = $false,

        # CategoryName - specify a single or multiple categories categorize the app (optional)
        [Parameter(Mandatory = $false)]
        [string[]]$CategoryName,

        # RestartBehavior - restart behavior if app requires restart (optional, defaults to basedOnReturnCode)
        [Parameter(Mandatory = $false)]
        [ValidateSet("allow", "basedOnReturnCode", "suppress", "force")]
        [string]$RestartBehavior = "basedOnReturnCode",

        # InstallExperience - install as system or user (optional, default SYSTEM)
        [Parameter(Mandatory = $false)]
        [ValidateSet("SYSTEM", "User")]
        [string]$InstallExperience = "SYSTEM",

        # InstallCommandLine - install command line (optional, defaults to psadt install)
        [Parameter(Mandatory = $false)]
        [string]$InstallCommandLine = ($SourcePath | Split-Path -Leaf) + " -DeploymentType Install",

        # UninstallCommandLine - uninstall command line (optional, defaults to psadt uninstall)
        [Parameter(Mandatory = $false)]
        [string]$UninstallCommandLine = ($SourcePath | Split-Path -Leaf) + " -DeploymentType Uninstall",

        # MaximumInstallationTimeInMinutes - limit the time installation can take before erroring (optional)
        [Parameter(Mandatory = $false)]
        [int]$MaximumInstallationTimeInMinutes,

        # AppVersion - version that appears in company portal (optional, defaults to registry version)
        [Parameter(Mandatory = $false)]
        [string]$AppVersion = $RegKeyVersionValue
    )

    $ErrorActionPreference = "Stop"

    # package folder containing installer files as .intunewin file
    $intuneWinPath = New-IntuneAppPackage -InstallerPath $SourcePath

    # get minimum disk space estimate
    $sourceFolder = $SourcePath | Split-Path -Parent
    $getMinDiskSpaceArgs = @{
        InstalledApplicationSizeMB  = $InstalledApplicationSizeMB
        SourceFolder                = $SourceFolder
        IntuneWinPath               = $intuneWinPath
    }
    $minDiskSpaceEstimate = Get-MinimumDiskSpaceEstimate @getMinDiskSpaceArgs

    # create requirement rule for architecture, windows version and disk space
    $newReqRuleArgs = @{
        Architecture                    = "AllWithARM64"
        MinimumSupportedWindowsRelease  = "W11_22H2"
        MinimumFreeDiskSpaceInMB        = $minDiskSpaceEstimate
    }
    $reqRule = New-IntuneWin32AppRequirementRule @newReqRuleArgs

    # Create registry version detection rule
    # if version has only numbers and periods then use greater than or equal
    # else use string comparison equal
    if ($RegKeyVersionValue -match "^(\d+(\.\d+){0,3})$") {
        $newDetRuleRegArgs = @{
            KeyPath                     = $RegKeyPath
            ValueName                   = $RegKeyValueName
            VersionComparisonValue      = $RegKeyVersionValue
            VersionComparison           = $true
            VersionComparisonOperator   = "greaterThanOrEqual"
        }
    } elseif (-not $RegKeyVersionValue) {
        $newDetRuleRegArgs = @{
            KeyPath                     = $RegKeyPath
            ValueName                   = $RegKeyValueName
            Existence                   = $true
            DetectionType               = "exists"
        }
    } else {
        $newDetRuleRegArgs = @{
            KeyPath                     = $RegKeyPath
            ValueName                   = $RegKeyValueName
            StringComparisonValue       = $RegKeyVersionValue
            StringComparison            = $true
            StringComparisonOperator    = "equal"
        }
    }

    $detRuleReg = New-IntuneWin32AppDetectionRuleRegistry @newDetRuleRegArgs

    # put all info we have so far into a hashtable
    # we'll add other properties if they have been provided
    $addIntuneWin32AppArgs = @{
        FilePath                    = $intuneWinPath
        DisplayName                 = $DisplayName
        Description                 = $Description
        Publisher                   = $Publisher
        AppVersion                  = $AppVersion
        Developer                   = $Developer
        CompanyPortalFeaturedApp    = $CompanyPortalFeaturedApp
        InstallExperience           = $InstallExperience
        InstallCommandLine          = $InstallCommandLine
        UninstallCommandLine        = $UninstallCommandLine
        RestartBehavior             = $RestartBehavior
        RequirementRule             = $reqRule
        AllowAvailableUninstall     = $true
        UseAzCopy                   = $true
        AzCopyWindowStyle           = "Hidden"
        ScopeTagName                = @(
            "Dept-001-H07-Corporate-IT",
            "Endpoint Team-au_it_level_2_sccm_access_usg",
            "OU-au_computers_sccm_2012"
        )
    }

    if ($PSBoundParameters["ApplicationLauncherPath"]) {
        # create file detection rule - checking for file's existence
        $appLauncherFolder = $ApplicationLauncherPath | Split-Path -Parent
        $appLauncherFile = $ApplicationLauncherPath | Split-Path -Leaf

        $newDetRuleFileArgs = @{
            Path            = $appLauncherFolder
            FileOrFolder    = $appLauncherFile
            Existence       = $true
            DetectionType   = "exists"
        }
        $detRuleFile = New-IntuneWin32AppDetectionRuleFile @newDetRuleFileArgs

        $addIntuneWin32AppArgs.Add("DetectionRule", @($detRuleFile, $detRuleReg))
    } else {
        $addIntuneWin32AppArgs.Add("DetectionRule", @($detRuleReg))
    }

    if ($PSBoundParameters["Owner"]) {
        $addIntuneWin32AppArgs.Add("Owner", $Owner)
    }

    if ($PSBoundParameters["Notes"]) {
        $addIntuneWin32AppArgs.Add("Notes", $Notes)
    }

    # Convert image file to base64 string
    if ($PSBoundParameters["IconPath"]) {
        $icon = New-IntuneWin32AppIcon -FilePath $IconPath

        $addIntuneWin32AppArgs.Add("Icon", $icon)
    }

    if ($PSBoundParameters["CategoryName"]) {
        $addIntuneWin32AppArgs.Add("CategoryName", $CategoryName)
    }

    if ($PSBoundParameters["InformationURL"]) {
        $addIntuneWin32AppArgs.Add("InformationURL", $InformationURL)
    }

    if ($PSBoundParameters["PrivacyURL"]) {
        $addIntuneWin32AppArgs.Add("PrivacyURL", $PrivacyURL)
    }

    if ($PSBoundParameters["MaximumInstallationTimeInMinutes"]) {
        $addIntuneWin32AppArgs.Add("MaximumInstallationTimeInMinutes", $MaximumInstallationTimeInMinutes)
    }

    if ($PSCmdlet.ShouldProcess($DisplayName, "Adding this app to Intune.")) {
        $publishedApp = Add-IntuneWin32App @addIntuneWin32AppArgs
        Write-Output $publishedApp.id
    }

    Write-Output $intuneWinPath
}
