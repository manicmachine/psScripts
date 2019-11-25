#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Date: 11/8/2019
# Description: Monitors a collection of computers to determine if they are available
# to be remotely accessed. If so, the script will alert the user which invoked the script.
#####
param(
    [string] $computer,
    [string[]] $computers = @(),
    [string] $filePath,
    [String] $sendEmail = "false"
)

Import-Module ActiveDirectory

# MonitorQueue class definition.
# Manages the monitor queue, checks for credentials, connectivity, availability,
# as well as generating alerts and reports.
class MonitorQueue {
    [System.Collections.ArrayList] $computers
    [System.Collections.ArrayList] $toBeRemoved
    [Hashtable] $credentials
    [Hashtable] $reports
    [String] $username
    [String] $time
    [bool] $sendEmail

    MonitorQueue([String] $username, [bool] $sendEmail) {
        $this.computers = New-Object System.Collections.ArrayList
        $this.toBeRemoved = New-Object System.Collections.ArrayList
        $this.credentials = @{ }
        $this.reports = [ordered] @{ }
        $this.username = $username
        $this.updateTime()
        $this.sendEmail = $sendEmail
    }

    addToQueue([string] $computer) {
        $this.computers.Add($computer)
        Write-Host "Added " -NoNewline
        Write-Host $computer -ForegroundColor Green -NoNewline
        Write-Host " to the queue."
    }

    removeFromQueue([string] $computer) {
        $this.computers.Remove($computer)
        Write-Host "Removed " -NoNewline
        Write-Host $computer -ForegroundColor Green -NoNewline
        Write-Host " from the queue."
    }

    updateQueue() {
        Foreach ($record in $this.toBeRemoved) {
            $this.removeFromQueue($record)
        }

        $this.toBeRemoved.Clear()
    }

    # Check if computer is reachable.
    [bool] checkConnectivity([string] $computer) {
        If (Test-Connection $computer -Count 3 -Quiet) {
            Return $true
        }
        Else {
            Return $false
        }
    }

    # Check if computer is in use.
    [bool] checkAvailability([string] $computer, [Hashtable] $credentials) {
        If ([String]::IsNullorEmpty($this.usedBy($computer, $credentials))) {
            Return $true
        }
        Else {
            Return $false
        }
    }

    # Generate alert that computer is available.
    alert([string] $computer) {
        [console]::beep(1700, 300)

        If ($this.sendEmail) {
            Send-MailMessage -From "#REDACTED#" -To "$($this.username)@uwec.edu" -Subject "$computer - User Logged Off" -Body "The user that was logged into $computer logged off at $($this.time)." -SmtpServer "#REDACTED#" -Port "#REDACTED#"
        }
    }

    # Prompt and store credentials for each type of computer (lab, wks, srv) provided.
    # NOTE: Credentials are stored securely and are only usable on the computer
    # which the credential file was created and by the user which created it.
    getCredentials([System.Collections.ArrayList] $computers) {
        Foreach ($computer in $computers) {
            $type = $this.computerType($computer)
            If (!($this.credentials.ContainsKey($type))) {
                If (Test-Path "$env:USERPROFILE\$($type).cred") {
                    $cred = Import-Clixml -Path "$env:USERPROFILE\$($type).cred"

                    $this.credentials.Add($type, $cred)
                    Break
                }
                Else {
                    Do {
                        $cred = Get-Credential -Credential "UWEC\$($this.username)-$type"
                        $tempUser = "$($this.username)-$type"
                        $password = $cred.GetNetworkCredential().password
                        $currentDomain = "LDAP://" + ([ADSI]"").distinguishedName
                        $domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain, $tempUser, $password)

                        If (!($domain.Name)) {
                            $badLoginPrompt = New-Object -ComObject Wscript.Shell
                            $badLogin = $badLoginPrompt.Popup("Authentication failed - would you like to try to enter credentials again?", 30, "Authentication Failed", 0x4)

                            Switch ($badLogin) {
                                '6' {
                                    Continue
                                }'7' {
                                    Exit
                                }'-1' {
                                    Exit
                                }
                            }
                        }
                        Else {
                            $cred | Export-CliXml -Path "$env:USERPROFILE\$($type).cred"
                            $this.credentials.Add($type, $cred)
                            Break
                        }
                    } While ($true)
                }
            }
        }
    }


    # Determine what type of computer (workstation, lab, server) is provided.
    [string] computerType([string] $computer) {
        $record = Get-ADComputer $computer -Properties DistinguishedName

        If ($record.DistinguishedName -match "ou=labs") {
            Return "lab"
        }
        Elseif ($record.DistinguishedName -match "ou=uwec servers") {
            Return "srv"
        }
        Else {
            Return "wks"
        }
    }

    # Determine who's currently using the specified computer.
    [string] usedBy([string] $computer, [Hashtable] $credentials) {
        $creds = $credentials[$this.computerType($computer)]
        $user = Get-WmiObject -Credential $creds –ComputerName $computer –Class Win32_ComputerSystem | Select-Object UserName
        Return $user.Username
    }

    # Add the most recent info for the computer into the reports.
    addReport([string] $computer, [string] $details) {
        If (!($this.reports.ContainsKey($computer))) {
            $this.reports.Add($computer, $details)
        }
        Else {
            $this.reports[$computer] = $details
        }
    }

    # Present the report to the user.
    printReport() {
        $this.updateTime()
        Clear-Host
        Write-Host "----------------------------------" -ForegroundColor Green
        Write-Host "Report as of $($this.time)"
        Write-Host "----------------------------------" -ForegroundColor Green

        Foreach ($computer in $this.reports.keys) {
            Write-Host "$($computer)" -NoNewline -ForegroundColor Green
            Write-Host " - " -NoNewline
            If ($this.reports[$computer] -match "Available!$") {
                Write-Host $this.reports[$computer] -ForegroundColor Green
            }
            Else {
                Write-Host $this.reports[$computer]
            }
        }

        Write-Host "----------------------------------" -ForegroundColor Green
        Write-Host "Computers left in queue: " $this.computers.Count
        Write-Host "----------------------------------" -ForegroundColor Green
    }

    updateTime() {
        $this.time = Get-Date -Format G
    }

    # Begin the monitoring queue.
    start() {
        # Make sure the provided computers are reachable. By using runspaces, the script
        # can check the connectivity of all computers at once without waiting
        # for individual timeouts sequentially.
        $computerConnections = @{ }
        $pool = [runspacefactory]::CreateRunspacePool(1, [int]$env:NUMBER_OF_PROCESSORS * 5)
        $pool.ApartmentState = "MTA" # Set RunSpacePool to be Multi-threaded
        $pool.Open()
        $runspaces = @()
        $connectionScriptBlock = {
            param (
                $computer)

            $reachable = Test-Connection -Count 3 $computer -Quiet
            return [PSCustomObject]@{ Computer = $computer; Reachable = $reachable }
        }

        Write-Host "Checking connectivity of computers in queue..."
        Foreach ($computer in $this.computers) {
            $runspace = [powershell]::Create()
            $runspace.AddScript($connectionScriptBlock)
            $runspace.AddArgument($computer)
            $runspace.RunspacePool = $pool

            $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
        }

        # Wait for runspaces to finish
        While ($runspaces.Status.IsCompleted -notcontains $true) {}

        # Grab results
        Foreach ($runspace in $runspaces) {
            $result = $runspace.Pipe.EndInvoke($runspace.Status)
            $computerConnections.add($result.Computer, $result.Reachable)
            $runspace.Pipe.Dispose()
        }

        Foreach ($computer in $this.computers) {
            If ($computerConnections[$computer]) {
                Write-Host "$computer" -NoNewline -ForegroundColor Green
                Write-Host " - $($this.time): Computer is reachable."
            }
            Else {
                $this.updateTime()
                Write-Host "$computer" -NoNewline -ForegroundColor Green
                Write-Host " - $($this.time): Computer is unreachable."
                $this.addReport($computer, "$($this.time): Unreachable.")
                $this.toBeRemoved.Add($computer)
            }
        }

        $pool.Close()
        $pool.Dispose()

        $this.updateQueue()
        $this.getCredentials($this.computers)

        # Begin monitoring loop.
        Do {

            $this.printReport()

            Foreach ($computer in $this.computers) {
                If (!($this.checkConnectivity($computer))) {
                    $this.updateTime()
                    Write-Host "$computer" -NoNewline -ForegroundColor Green
                    Write-Host " - $($this.time): Computer is unreachable."
                    $this.addReport($computer, "$($this.time): Unreachable.")
                    $this.toBeRemoved.Add($computer)
                    continue
                }

                If ($this.checkAvailability($computer, $this.credentials)) {
                    $this.updateTime()
                    Write-Host "$computer" -NoNewline -ForegroundColor Green
                    Write-host " - $($this.time): Available!"
                    $this.addReport($computer, "$($this.time): Available!")
                    $this.alert($computer)
                    $this.toBeRemoved.Add($computer)
                }
                Else {
                    $this.updateTime()
                    $user = $this.usedBy($computer, $this.credentials)
                    Write-Host "$computer" -NoNewline -ForegroundColor Green
                    Write-host " - $($this.time): Currently in-use by $user."
                    $this.addReport($computer, "$($this.time): Unavailable; In-use by $user")
                }
            }

            $this.updateQueue()
            If ($this.computers.Count -gt 0) {
                Write-Host "End of queue. Waiting 30 seconds..."
                Start-Sleep -Seconds 30
            }

        } While ($this.computers.Count -gt 0)

        # All computers are accounted for. Print final report.
        $this.printReport()
    }
}

# Check if the provided computer is available within AD.
function isValid([string] $computer) {
    Try {
        $validComputer = Get-ADComputer $computer
        If ($validComputer) { Return $true }
    }
    Catch {
        Return $false
    }
}

# Initialize the necessary variables.
$username = If ($env:username -match "(?<username>.+).*-lab|(?<username>.+).*-srv|(?<username>.+).*-wks") {
    $matches["username"]
}
Else {
    $env:username
}
$monitorQueue = New-Object MonitorQueue($username, [System.Convert]::ToBoolean($sendEmail))

# If nothing is passed to the script, prompt for a computer name or file path.
If (!($computer -or $computers.Count -ne 0 -or $filePath)) {
    Write-Host "Enter one of the following:"
    Write-host "- A computer name"
    Write-Host "- A list of computer names, comma seperated"
    Write-Host "- The path to a file containing a list of computers"
    $hostInput = Read-Host "Input"

    If (Test-Path $hostInput) {
        $filePath = $hostInput
    }
    Elseif ($hostInput -match ",") {
        $computers = $hostInput.Split(",").Trim()
    }
    Else {
        $computer = $hostInput
    }
}

# If a path is provided, and valid, read contents and populate computers array.
# Added logic to accomidate comma seperated lists and those with extra whitespace/newlines.
If (($filePath) -and (Test-Path $filePath)) {
    Foreach ($computer in (Get-Content -Path $filePath).Split(",")) {
        If (!([String]::IsNullorEmpty($computer.Trim()))) {
            $computers += $computer
        }
    }

    # Post-loop cleanup. Necessary due to shared variable name and scope weirdness.
    Clear-Variable computer
}

# Begin!
$monitorQueue.printReport() # Clear screen and present report dashboard.
If ($computer) { $computers += $computer }
Foreach ($computer in $computers) {
    If (isValid $computer) {
        $monitorQueue.addToQueue($computer.toUpper())
    }
    Else {
        Write-Host "$computer was not found. Ignored."
    }
}

If ($monitorQueue.computers.Count -ne 0) {
    $monitorQueue.start()
}
Else {
    Write-Host "No computers found in the queue. Exiting."
}