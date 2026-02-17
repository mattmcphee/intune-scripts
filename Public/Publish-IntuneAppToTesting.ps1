function Publish-IntuneAppToTesting {
    [CmdletBinding(DefaultParameterSetName = "ByDisplayName")]
    param (
        # DisplayName - displayname of app in intune
        [Parameter(Mandatory=$true, ParameterSetName="ByDisplayName")]
        [string]$DisplayName,

        # AppID - ID of application
        [Parameter(Mandatory=$true, ParameterSetName="ByAppID")]
        [string]$AppID
    )

    if ($PSCmdlet.ParameterSetName -eq "ByDisplayName") {
        try {
            $win32App = Get-IntuneWin32App -DisplayName $DisplayName
        } catch {
            throw "Encountered error when retrieving app from intune with displayname: '$DisplayName'"
        }
    
        if ($null -eq $win32App) {
            throw "Could not find app with DisplayName: '$DisplayName'"
        } elseif ($win32App.Count -gt 1) {
            throw "Found more than 1 app with DisplayName: '$DisplayName'. Be more specific."
        }
    }

    if ($PSCmdlet.ParameterSetName -eq "ByAppID") {
        $id = $AppID
    } else {
        $id = $win32App.id
    }

    try {
        $intuneWin32AppAssignmentArgs = @{
            Include         = $true
            ID              = $id
            GroupID         = "9deb0e98-7051-4279-9e7a-31ec1daae2b9"
            Intent          = "available"
            Notification    = "showAll"
        }

        Add-IntuneWin32AppAssignmentGroup @intuneWin32AppAssignmentArgs
    } catch {
        throw "Encountered error when assigning app with id '$id' to mem-device-app-testing_gs with group ID '9deb0e98-7051-4279-9e7a-31ec1daae2b9'"
    }
}
