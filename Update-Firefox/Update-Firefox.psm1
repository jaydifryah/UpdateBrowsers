<#
.SYNOPSIS
    Downloads and installs the latest extended support release of Firefox from Mozilla
.DESCRIPTION
    Downloads the latest Firefox ESR, copies it to system temp, and installs
    Downloads the file from https://download.mozilla.org/?product=Firefox-esr-latest-ssl&os=win64&lang=en-US
.EXAMPLE
    Update-Firefox
        Downloads and installs latest ESR.  Will prompt for hostname and assume a single target
.EXAMPLE
    Update-Firefox -ComputerName hostname
        Downloads and installs latest ESR on given host
.EXAMPLE
    Update-Firefox -File .\targets.txt
        Downloads and installs latest ESR on all listed targets, returns a table of results
.INPUTS
    Single hostname [-ComputerName]
    Text file of host names [-File]
.OUTPUTS
    Table of:
    - Computer names
    - Firefox version prior to script running
    - Firefox version of update installer downloaded
    - Firefox version after script running
    - If an update was applied for given targets
.NOTES
    Author:         jaydifryah
    Creation Date:  04/2020
#>
function Update-Firefox {
    [CmdletBinding(DefaultParameterSetName="ComputerName")]
    param (
        # Used for passing in a single host
        [Parameter(
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Mandatory=$true,
            ParameterSetName="ComputerName",
            HelpMessage="Enter a single host name here"
        )]
        [String]
        $ComputerName,

        # Used for passing in a text file with multiple hosts
        [Parameter(
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Mandatory=$true,
            ParameterSetName="File",
            HelpMessage="Enter a text file with multiple host names here"
        )] [ValidateScript(
            { Test-Path $_ -PathType 'Leaf' }
        )] [ValidateScript(
            { (Get-Item $_).Extension -eq ".txt" }
        )]
        [String]
        $File
    )

    begin {
        # Setting variables to check which parameter was selected
        $cnParam = ($PSBoundParameters.ContainsKey('ComputerName'))
        $fileParam = ($PSBoundParameters.ContainsKey('File'))

        # Setting variables for downloaded file, destination, current exe location, and version check
        $uri = "https://download.mozilla.org/?product=Firefox-esr-latest-ssl&os=win64&lang=en-US"
        $workingFolder = "$env:windir\temp\FirefoxESR"
        $destination = "$env:windir\temp\FirefoxESR.exe"
        $ffExe = '"$env:ProgramFiles\Mozilla Firefox\Firefox.exe"'
        $global:ffVer = "(Get-Item -Path $ffExe).VersionInfo.FileVersion"

        # Remove old update file if present (local)
        if (Test-Path $workingFolder) {
            Remove-Item $workingFolder -Recurse -Force | Out-Null
        } elseif (Test-Path $destination) {
            Remove-Item $destination -Recurse -Force | Out-Null
        }

        # Get a copy of Firefox locally, extract it to perform version checks
        # Initialize Web Client to perform download, download Firefox, save it to $destination
        New-Item -Path $workingFolder -ItemType Directory -Force | Out-Null
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($uri, $destination)

        # Extract installer locally to check version
        Start-Process -FilePath $destination -ArgumentList "/ExtractDir=$workingFolder" -NoNewWindow -Wait
        $firefoxUpdateVersion = (Get-Item -Path "$workingFolder\core\Firefox.exe").VersionInfo.FileVersion

        $scriptBlock = {
            # Remove old update file if present (remote)
            if (Test-Path $using:destination) {
                Remove-Item $using:destination -Force | Out-Null
            }

            # Get current version of Firefox before install
            $firefoxOldVersion = Invoke-Expression $using:ffVer

            # Check current version against installer version, skip if already updated
            if ($firefoxOldVersion -lt $using:firefoxUpdateVersion) {
                # New Web Client needs to be initialized on remote machine
                # Initialize Web Client, download Firefox, save it to $destination
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($using:uri, $using:destination)

                # Start install
                Start-Process -FilePath $using:destination -ArgumentList "/S" -NoNewWindow -Wait

                # Set 'Installed' flag
                $firefoxInstalled = $true

                # Check if Firefox is running on remote, indicate need for restart if so
                $firefoxRunning = (Get-Process -Name "*firefox*").count -ne 0

                # Check version post-install
                $firefoxCurrentVersion = Invoke-Expression $using:ffVer

            } elseif ($firefoxOldVersion -eq $using:firefoxUpdateVersion) {
                # Set 'Current' flag
                $firefoxCurrent = $true
                $firefoxCurrentVersion = Invoke-Expression $using:ffVer
            }

            # Get installer version, newly installed version for verification
            $props = [ordered]@{
                ComputerName = "$env:COMPUTERNAME"
                Firefox_Old_Version = $firefoxOldVersion
                Installer_Version = $using:firefoxUpdateVersion
                Firefox_New_Version = $firefoxCurrentVersion
            }

            # 'Updated' column logic
            # Check if version was already current
            if ($firefoxCurrent -eq $true) {
                $props.Add('Updated','Current')
            }

            # Otherwise check if Firefox did install, but the
            # version didn't change, due to it running
            # If the version didn't change, but Firefox isn't
            # running, the install was not sucessful
            elseif ($firefoxInstalled -eq $true) {
                if ($firefoxCurrentVersion -eq $using:firefoxUpdateVersion) {
                    $props.Add('Updated',$true)
                } elseif ($firefoxOldVersion -eq $firefoxCurrentVersion) {
                    if ($firefoxRunning -eq $true) {
                        $props.Add('Updated','Needs Restart')
                    } else {
                        $props.Add('Updated',$false)
                    }
                }
            } else {
                if ($firefoxCurrentVersion -eq $using:firefoxUpdateVersion) {
                    $props.Add('Updated',"Current")
                } else {
                    $props.Add('Updated',$false)
                }
            }

            # Generate result object
            New-Object -TypeName psobject -Property $props

            # Remove installer
            Remove-Item -Path $workingFolder -Recurse -Force | Out-Null
        }

    }

    process {
        # Check which parameter was used, execute against appropriate target
        if ($cnParam) {
            $cnResults = Invoke-Command -ComputerName $ComputerName -ScriptBlock `
                $scriptBlock -ErrorAction "SilentlyContinue"
        } elseif ($fileParam) {
            $txtFile = Get-Content -Path $File
            $fileResults = Invoke-Command -ComputerName $txtFile -ScriptBlock `
                $scriptBlock -ThrottleLimit 16 -ErrorAction "SilentlyContinue"
        }
    }

    end {
        # Provide results
        if ($cnParam) {
            $cnResults | Select-Object -Property `
                ComputerName,Firefox_Old_Version,Installer_Version,Firefox_New_Version,Updated | Format-Table -AutoSize
        } elseif ($fileParam) {
            $fileResults | Select-Object -Property `
                ComputerName,Firefox_Old_Version,Installer_Version,Firefox_New_Version,Updated | Format-Table -AutoSize
        }
    }
}

# Define functions to export
Export-ModuleMember -Function "Update-Firefox"