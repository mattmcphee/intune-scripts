function Get-MinimumDiskSpaceEstimate {
    [CmdletBinding()]
    [OutputType([int])]
    param (
        # InstalledApplicationSizeMB - number of MB the application takes up after installation
        [Parameter(Mandatory=$true)]
        [int]$InstalledApplicationSizeMB,

        # SourceFolder - path to folder containing installer files
        [Parameter(Mandatory=$true)]
        [ValidateScript({
            if (Test-Path $_ -PathType Container) {
                return $true
            } else {
                throw "SourceFolder parameter '$_' is not a valid folder path."
            }
        })]
        [string]$SourceFolder,

        # IntuneWinPath - path to .intunewin file containing compressed installer files
        [Parameter(Mandatory=$true)]
        [ValidateScript({
            if (Test-Path $_ -PathType Leaf) {
                return $true
            } else {
                throw "IntuneWinPath parameter '$_' is not a valid .intunewin file path."
            }
        })]
        [string]$IntuneWinPath
    )

    # calculate how much disk space is required
    # add 1.5x the installed app size as overhead
    # add 100MB on top just in case
    try {
        [int]$sourceFolderSizeMB = [int]((Get-ChildItem -Path $SourceFolder -Recurse -Force | Measure-Object -Property Length -Sum).Sum / 1MB)
        [int]$intunewinSizeMB = [int]((Get-ChildItem -Path $IntuneWinPath -Force | Measure-Object -Property Length -Sum).Sum / 1MB)
        [int]$installedAppSizeWithOverhead = [int]($InstalledApplicationSizeMB * 1.5)

        [int]$minFreeDiskSpaceMB = $sourceFolderSizeMB + $intunewinSizeMB + $installedAppSizeWithOverhead + 100
        # round up to the nearest 100MB just in case in case
        [int]$minFreeDiskSpaceMB = $minFreeDiskSpaceMB + (100 - ($minFreeDiskSpaceMB % 100))

        return $minFreeDiskSpaceMB
    } catch {
        throw "Error encountered when calculating minimum free disk space: $_"
    }
}
