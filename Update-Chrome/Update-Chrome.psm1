<#
.SYNOPSIS
    Downloads and installs the latest extended support release of Chrome from Google
    Requires Install-Package MSI -Provider PowerShellget
.DESCRIPTION
    Downloads the latest Google Chrome enterprise release, copies it to system temp, and installs
    Downloads the file from https://dl.google.com/tag/s/dl/chrome/install/googlechromestandaloneenterprise64.msi
.EXAMPLE
    DPS-UpdateChrome
        Downloads and installs latest enterprise release.  Will prompt for hostname and assume a single target
.EXAMPLE
    Update-Chrome -ComputerName hostname
        Downloads and installs latest enterprise release on given host
.EXAMPLE
    Update-Chrome -File .\targets.txt
        Downloads and installs latest enterprise release on all listed targets, returns a table of results
.INPUTS
    Single hostname [-ComputerName]
    Text file of host names [-File]
.OUTPUTS
    Table of:
    - Computer names
    - Chrome version prior to script running
    - Chrome version of update installer downloaded
    - Chrome version after script running
    - If an update was applied for given targets
.NOTES
    Author:         jaydifryah
    Creation Date:  04/2020
#>
function Update-Chrome {
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

        # Setting variables for downloaded file, destination, current exe location and version
        $uri = "https://dl.google.com/tag/s/dl/chrome/install/googlechromestandaloneenterprise64.msi"
        $destination = "$env:windir\temp\ChromeEnt.msi"
        $chromeExe = '"${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"'
        $global:chromeVer = "(Get-Item -Path $chromeExe).VersionInfo.FileVersion"

        # Check for MSI info library dependency
        if ( (Get-Package -Name MSI -ErrorAction "SilentlyContinue").count -ne 0 ) {

            # Get a copy of Chrome locally to perform version checks
            # Initialize Web Client to perform download, download Chrome, save it to $destination
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($uri, $destination)

            # get Chrome version from MSI for verification
            $chromeMsiVersion = (Get-MSISummaryInfo $destination).Comments.Split(" ")[0]
            $nolibs = $false
        } else {
            Write-Warning "Please run 'Install-Package MSI -Provider PowerShellGet' for current version information"
            $chromeMsiVersion = "Unknown"
            $nolibs = $true
        }

        $scriptBlock = {
            # Remove old update file if present
            if (Test-Path $using:destination) {
                Remove-Item $using:destination -Force | Out-Null
            }

            # Get current version of Chrome before install
            $chromeOldVersion = Invoke-Expression $using:chromeVer

            # Check current version against MSI reference version, skip if already updated
            if ($chromeOldVersion -lt $using:chromeMsiVersion) {
                # New Web Client needs to be initialized on remote machine
                # Initialize Web Client, download Chrome, save it to $destination
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($using:uri, $using:destination)

                # MSI install parameters
                    $msiInstall = @(
                        "/i",
                        "$using:destination",
                        "/quiet",
                        "/qn",
                        "/norestart"
                    )

                # Start install
                Start-Process -FilePath "$env:windir\system32\msiexec.exe" -ArgumentList $msiInstall -Wait -NoNewWindow

                # Set 'Installed' flag
                $chromeInstalled = $true

                # Check if Chrome is running on remote, indicate need for restart if so
                $chromeRunning = (Get-Process -Name "*chrome*").count -ne 0

                # Check version post-install
                $chromeCurrentVersion = Invoke-Expression $using:chromeVer

            } elseif ($chromeOldVersion -eq $using:chromeMsiVersion) {
                # Set 'Current' flag
                $chromeCurrent = $true
                $chromeCurrentVersion = Invoke-Expression $using:chromeVer
            }

            # Get MSI file version, newly installed version for verification
            $props = [ordered]@{
                ComputerName = $env:COMPUTERNAME
                Chrome_Old_Version = $chromeOldVersion
                Installer_Version = $using:chromeMsiVersion
                Chrome_Current_Version = $chromeCurrentVersion
            }

            # 'Updated' column logic
            # If no MSI libraries, no update version to check
            if ($using:noLibs -eq $true) {
                $props.Add('Updated','Unknown')
            }

            # Otherwise check if version was already current
            elseif ($chromeCurrent -eq $true) {
                $props.Add('Updated','Current')
            }

            # Otherwise check if Chrome did install, but the
            # version didn't change, due to it running
            # If the version didn't change, but Chrome isn't
            # running, the install was not sucessful
            elseif ($chromeInstalled -eq $true) {
                if ($chromeCurrentVersion -eq $using:chromeMsiVersion) {
                    $props.Add('Updated',$true)
                } elseif ($chromeOldVersion -eq $chromeCurrentVersion) {
                    if ($chromeRunning -eq $true) {
                        $props.Add('Updated','Needs Restart')
                    } else {
                        $props.Add('Updated',$false)
                    }
                }
            } else {
                if ($chromeCurrentVersion -eq $using:chromeMsiVersion) {
                    $props.Add('Updated',"Current")
                } else {
                    $props.Add('Updated',$false)
                }
            }

            # Generate result object
            New-Object -TypeName psobject -Property $props

            # Remove installer
            Remove-Item -Path $using:destination -Force | Out-Null
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
                ComputerName,Chrome_Old_Version,Installer_Version,Chrome_Current_Version,Updated | Format-Table -AutoSize
        } elseif ($fileParam) {
            $fileResults | Select-Object -Property `
                ComputerName,Chrome_Old_Version,Installer_Version,Chrome_Current_Version,Updated | Format-Table -AutoSize
        }

    }
}

# Define functions to export
Export-ModuleMember -Function "DPS-UpdateChrome"