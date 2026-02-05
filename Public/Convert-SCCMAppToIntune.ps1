function Convert-SCCMAppToIntune {
    [CmdletBinding()]
    param (
        # ApplicationName - exact name as it appears in SCCM
        [Parameter(Mandatory)]
        [string]$ApplicationName,

        # IconPath - path to image icon
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

        # InstalledApplicationSizeMB
        [Parameter(Mandatory)]
        [int]$InstalledApplicationSizeMB,

        # Developer - author of the application (optional)
        [Parameter(Mandatory = $false)]
        [string]$Developer,

        # Owner - product owner of the application within the organization (optional)
        [Parameter(Mandatory = $false)]
        [string]$Owner,

        # Notes - optional
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

        # CategoryName - specify a single or multiple categories to categorize the app (optional)
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

        # InstallCommandLine - install command line (optional)
        [Parameter(Mandatory = $false)]
        [string]$InstallCommandLine,

        # UninstallCommandLine - uninstall command line (optional)
        [Parameter(Mandatory = $false)]
        [string]$UninstallCommandLine,

        # ApplicationExePath - path to main exe on the client machine (will be used for detection if provided)
        [Parameter(Mandatory = $false)]
        [string]$ApplicationExePath,

        # RegKeyPath - registry path that holds displayversion value name on the client machine
        # (will be used for detection if provided)
        # Must include HKLM\ or HKEY_LOCAL_MACHINE\ at the start (or whatever registry hive)
        [Parameter(Mandatory = $false)]
        [string]$RegKeyPath,

        # DisplayVersion - the version value that appears in the registry on the client machine (will be used for detection if provided)
        [Parameter(Mandatory = $false)]
        [string]$DisplayVersion
    )

    # get app
    try {
        $ogLoc = Get-Location
        Set-Location "A00:"
        $app = Get-CMApplication -Name $ApplicationName
        Set-Location $ogLoc
    } catch {
        throw "Could not find application with that name: $_"
    }

    # pull out info from sdmpackagexml
    [xml]$appxml = $app.SDMPackageXml
    $appxmlDisplayInfo = $appxml.AppMgmtDigest.Application.DisplayInfo.Info
    $appXmlCustomData = $appxml.AppMgmtDigest.DeploymentType.Installer.CustomData

    $displayName = $appxmlDisplayInfo.Title
    $description = $appxmlDisplayInfo.Description
    $publisher = $appxmlDisplayInfo.Publisher
    $appVersion = $appxmlDisplayInfo.Version

    # base folder and installer file are two separate elements in xml
    $installerFile = $appXmlCustomData.InstallCommandLine -split " " |
        Select-Object -First 1
    $sourcePathBase = $appxml.AppMgmtDigest.DeploymentType.Installer.Contents.Content |
        Where-Object { $_.Location -notlike "*\Uninstall\*" } |
        Select-Object -ExpandProperty Location
    $sourcePath = $sourcePathBase + $installerFile

    # use exe path if provided or try to pull it out from sccm
    if ($PSBoundParameters["ApplicationExePath"]) {
        $applicationExePath = $ApplicationExePath
    } else {
        $appPathBase = $appXmlCustomData.EnhancedDetectionMethod.Settings.File.Path
        $appPathExe = $appXmlCustomData.EnhancedDetectionMethod.Settings.File.Filter

        if ([string]::IsNullOrEmpty($appPathBase) -or [string]::IsNullOrEmpty($appPathExe)) {
            $msg = "Unable to find file existence detection method for: $($app.LocalizedDisplayName)`n" +
            "Provide a path to the installed application's executable."
            throw $msg
        }

        $applicationExePath = "$appPathBase\$appPathExe"
    }

    # reg key path and reg key value are two separate elements
    if ($PSBoundParameters["RegKeyPath"] -and $PSBoundParameters["DisplayVersion"]) {
        $regKeyPath = $RegKeyPath
        $displayVersion = $DisplayVersion
    } else {
        $regKeyPath = "HKEY_LOCAL_MACHINE\" + $appXmlCustomData.EnhancedDetectionMethod.Settings.SimpleSetting.RegistryDiscoverySource.Key
        $displayVersion = $appXmlCustomData.EnhancedDetectionMethod.Rule.Expression.Operands.Expression.Operands.ConstantValue |
            Where-Object { $_.DataType -like "String" } |
            Select-Object -ExpandProperty Value
        if ([string]::IsNullOrEmpty($regKeyPath) -or [string]::IsNullOrEmpty($displayVersion)) {
            throw "Unable to find registry key detection method for: $($app.LocalizedDisplayName)"
        }
    }

    if ($PSBoundParameters["InstallCommandLine"]) {
        $installCommandLine = $InstallCommandLine
    } else {
        $installCommandLine = $appXmlCustomData.InstallCommandLine
    }

    if (-not $installCommandLine) {
        throw "Could not find install command line: $_"
    }

    if ($PSBoundParameters["UninstallCommandLine"]) {
        $uninstallCommandLine = $UninstallCommandLine
    } else {
        $uninstallCommandLine = $appXmlCustomData.UninstallCommandLine
    }

    if (-not $UninstallCommandLine) {
        throw "Could not find uninstall command line: $_"
    }

    # build splat
    $newIntuneAppArgs = @{
        SourcePath                  = $sourcePath
        InstalledApplicationSizeMB  = $InstalledApplicationSizeMB
        DisplayName                 = $displayName
        Publisher                   = $publisher
        Description                 = $description
        ApplicationExePath          = $applicationExePath
        RegKeyPath                  = $regKeyPath
        DisplayVersion              = $displayVersion
        AppVersion                  = $appVersion
        InstallCommandLine          = $installCommandLine
        UninstallCommandLine        = $uninstallCommandLine
    }

    if ($PSBoundParameters["IconPath"]) {
        $newIntuneAppArgs.Add("IconPath", $IconPath)
    }

    if ($PSBoundParameters["Developer"]) {
        $newIntuneAppArgs.Add("Developer", $Developer)
    }

    if ($PSBoundParameters["Owner"]) {
        $newIntuneAppArgs.Add("Owner", $Owner)
    }

    if ($PSBoundParameters["Notes"]) {
        $newIntuneAppArgs.Add("Notes", $Notes)
    }

    if ($PSBoundParameters["InformationURL"]) {
        $newIntuneAppArgs.Add("InformationURL", $InformationURL)
    }

    if ($PSBoundParameters["PrivacyURL"]) {
        $newIntuneAppArgs.Add("PrivacyURL", $PrivacyURL)
    }

    if ($PSBoundParameters["CompanyPortalFeaturedApp"]) {
        $newIntuneAppArgs.Add("CompanyPortalFeaturedApp", $CompanyPortalFeaturedApp)
    }

    if ($PSBoundParameters["RestartBehavior"]) {
        $newIntuneAppArgs.Add("RestartBehavior", $RestartBehavior)
    }

    if ($PSBoundParameters["InstallExperience"]) {
        $newIntuneAppArgs.Add("InstallExperience", $InstallExperience)
    }

    # create app
    try {
        $app = New-IntuneApp @newIntuneAppArgs
    } catch {
        throw "Error encountered when creating app in Intune: $_"
    }

    # assign to mem-device-app-testing_gs
    Publish-IntuneAppToTesting -AppID $app.id

    # output the parameter splat to the console
    Write-Output "`$convertSCCMAppArgs = @{"
    $PSBoundParameters.GetEnumerator() | ForEach-Object {
        Write-Output "$($_.Key) = `"$($_.Value)`""
    }
    Write-Output "}"
    Write-Output "Convert-SCCMAppToIntune @convertSCCMAppArgs"
}
