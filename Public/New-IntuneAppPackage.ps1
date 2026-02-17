function New-IntuneAppPackage {
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([string])]
    param (
        # InstallerPath - full path to installer exe e.g. deploy-application.exe or invoke-appdeploytoolkit.exe
        [Parameter(Mandatory=$true)]
        [ValidateScript({ 
            if ( Test-Path $_ -PathType Leaf ) {
                return $true
            } else {
                throw "InstallerPath parameter must be a valid path to an installer file."
            }
        })]
        [string[]]$InstallerPath
    )

    foreach ($installer in $InstallerPath) {
        $sourceFolder = $installer | Split-Path -Parent
        $setupFile = $installer | Split-Path -Leaf
        $outputFolder = $sourceFolder + " Intune"
        $setupFileNoExtension = [System.IO.Path]::GetFileNameWithoutExtension($setupFile)

        if (-not (Test-Path -Path $outputFolder -PathType Container)) {
            $null = New-Item -Path $outputFolder -ItemType Directory
        }

        if ($PSCmdlet.ShouldProcess("$sourceFolder", "Compressing $sourceFolder into $outputFolder\$setupFileNoExtension.intunewin")) {
            try {
                $null = New-IntuneWin32AppPackage -SourceFolder $sourceFolder -SetupFile $setupFile -OutputFolder $outputFolder
                Write-Verbose "Intune package exists at: $outputFolder\$setupFileNoExtension.intunewin"
                return "$outputFolder\$setupFileNoExtension.intunewin"
            } catch {
                throw "Error encountered when running New-IntuneWin32AppPackage: $_"
            }
        }
    }
}
