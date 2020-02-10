#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Date: 1/30/2020
# Description: Monitors a collection of computers to determine if they are available
# to be remotely accessed. If so, the script will alert the user which invoked the script.
#####
param(
    [string] $filePath = "C:\ComputerMonitor\computer_monitor_list.csv",
    [string] $logPath = "C:\ComputerMonitor\computer_monitor.log"
)

Import-Module ActiveDirectory

# MonitorQueue class definition.
# Manages the monitor queue, checks for connectivity, and generates alerts
class MonitorQueue {
    [System.Collections.ArrayList] $computers
    [System.Collections.ArrayList] $toBeRemoved
    [String] $time
    [String] $filePath
    [String] $logPath

    MonitorQueue([String] $filePath, [String] $logPath) {
        $this.computers = New-Object System.Collections.ArrayList
        $this.toBeRemoved = New-Object System.Collections.ArrayList
        $this.updateTime()
        $this.filePath = $filePath
        $this.logPath = $logPath
    }

    # Add computer to the monitor's queue.
    addToQueue([PSCustomObject] $computer) {
        $this.computers.Add($computer)
        $this.writeLog("Added $($computer.ComputerName) to the queue")
    }

    # Remove computer from the monitor's queue.
    removeFromQueue([PSCustomObject] $computer) {
        $this.computers.Remove($computer)
        $this.writeLog("Removed $($computer.ComputerName) from the queue")
    }

    # Update the monitor queue, adding or removing computers as necessary
    updateQueue() {
        $toBeAdded = @()
        $invalidNames = @()

        # Check that file isn't empty
        If (!([String]::IsNullOrEmpty($(Get-Content -Path $this.filePath)))) {

            $csv = Import-CSV -Path $this.filePath

            # Read file to see if any new computers have been added to the list
            Foreach ($computer in $csv) {
                # Add to queue if name is a valid computer and not already contained within the queue
                If (!($computer.Status -match "invalid") -and $this.isValid($computer)) {
                    If ($this.computers.Count -eq 0) {
                        $computer.Status = "Monitoring"
                        $toBeAdded += $computer
                    } Else {
                        If (!($this.computers.ComputerName.Contains($computer.ComputerName))) {
                            $computer.Status = "Monitoring"
                            $toBeAdded += $computer
                        }
                    }
                } ElseIf ($computer.Status -match "invalid" -or !($this.isValid($computer))) {
                    If (!($computer.Status -match "invalid")) {
                        $this.writeLog("$($computer.ComputerName) is an invalid computer name")
                        $computer.Status = "Invalid Computer Name"
                    }

                    $invalidNames += $computer
                }
            }

            # Add new computers to queue
            Foreach ($computer in $toBeAdded) {
                $this.addToQueue($computer)
            }

            # Remove computers from the queue which have been discovered online
            Foreach ($computer in $this.toBeRemoved) {
                $this.removeFromQueue($computer)
            }

            $this.toBeRemoved.Clear()
        }

        # Update computers file
        $output = [System.Text.StringBuilder]::new()
        $output.AppendLine("ComputerName, Comment, Status") # CSV Headers
        Foreach ($computer in $this.computers) {
            $output.AppendLine("$($computer.ComputerName), $($computer.Comment), $($computer.Status)")
        }

        Foreach ($computer in $invalidNames) {
            $output.AppendLine("$($computer.ComputerName), $($computer.Comment), $($computer.Status)")
        }

        # Write computers file
        $output.toString() | Out-File -FilePath $this.filePath
    }

    # Generate alert that computer is available.
    alert([PSCustomObject] $computer) {
        $subject = "$($computer.ComputerName) is now online"
        $body = [String]::Empty

        If ([String]::IsNullOrEmpty($computer.Status)) {
            $body = "$computer is online as of $($this.time)."
        } Else {
            $body = "$($computer.ComputerName) is online as of $($this.time). Monitoring Comment: $($computer.Comment)."
        }

        Send-MailMessage -From "#REDACTED#" -To "#REDACTED#" -Cc "#REDACTED#" -Subject $subject -Body $body -SmtpServer "#REDACTED#" -Port "2525"
    }

    # Append provided text to the log file
    writeLog([string] $logText) {
        $this.updateTime()
        Write-Output "$($this.time) - $logText" | Out-File -Append -FilePath $this.logPath
    }

    # Update the time using a readble format
    updateTime() {
        $this.time = Get-Date -Format G
    }

    # Check if the provided computer is available within AD.
    [Boolean]isValid([PSCustomObject] $computer) {
        Try {
            $validComputer = Get-ADComputer $($computer.ComputerName)
            If ($validComputer) { 
                Return $true 
            } Else {
                Return $false
            }
        } Catch {
            Return $false
        }
    }

    # Begin the monitoring queue.
    start() {
        # Make sure the provided computers are reachable. By using runspaces, the script
        # can check the connectivity of all computers at once without waiting
        # for individual timeouts sequentially.
        $sleepTimer = 1800 # Seconds
        $computerConnections = @{}
        $pool = [runspacefactory]::CreateRunspacePool(1, [int]$env:NUMBER_OF_PROCESSORS * 5)
        $pool.ApartmentState = "MTA" # Set RunSpacePool to be Multi-threaded
        $pool.Open()
        $runspaces = New-Object System.Collections.ArrayList
        $connectionScriptBlock = {
            param (
                [PSCustomObject] $computer)

            $fqdn = $computer.ComputerName

            # Make the computer name a fully-qualified domain name
            If ($fqdn.Substring(0, 4) -match "lab-") {
                $fqdn += ".labs.uwec.edu"
            } Else {
                $fqdn += ".offices.uwec.edu"
            }

            $reachable = Test-Connection -Count 3 $fqdn -Quiet
            return [PSCustomObject]@{ Computer = $computer; Reachable = $reachable }
        }

        Do {
            $this.writeLog("Beginning monitor run...")
            $this.updateQueue()

            If ($this.computers.Count -gt 0) {
                Foreach ($computer in $this.computers) {
                    $runspace = [powershell]::Create()
                    $runspace.AddScript($connectionScriptBlock)
                    $runspace.AddArgument($computer)
                    $runspace.RunspacePool = $pool

                    $runspaces.Add([PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() })
                }

                # Wait for runspaces to finish
                While ($runspaces.Status.IsCompleted -notcontains $true) {}

                # Grab results
                Foreach ($runspace in $runspaces) {
                    $result = $runspace.Pipe.EndInvoke($runspace.Status)
                    $computerConnections.add($result.Computer, $result.Reachable)
                    $runspace.Pipe.Dispose()
                }

                # Process results
                Foreach ($computer in $this.computers) {
                    If ($computerConnections[$computer]) {
                        $this.writeLog("$($computer.ComputerName) is reachable")
                        $this.alert($computer)
                        $this.toBeRemoved.Add($computer)
                    }
                }
            }

            $runspaces.clear()
            $computerConnections.Clear()
            $this.updateQueue()
            $this.writeLog("Monitor run complete")
            $this.writeLog("$($this.computers.Count) computer(s) currently in the monitor queue")

            Foreach ($computer in $this.computers) {
                $this.writeLog("`t $($computer.ComputerName)")
            }

            $this.writeLog("Waiting $($sleepTimer / 60) minutes...")
            Start-Sleep -Seconds $sleepTimer # Wait for next cycle
        } While ($true)
    }
}

$monitorQueue = New-Object MonitorQueue($filePath, $logPath)

# Check that computer list file exists
If (!(Test-Path $filePath)) {
    New-Item -Path $filePath -ItemType file -Force
}

# Check that log file isn't too big, if so, recreate it
If ((Test-Path -Path $logPath) -and $((Get-Item $logPath).Length)/1MB -gt 10) {
   Remove-Item -Path $logPath -Force
}

# Begin!
$monitorQueue.writeLog("New monitor run initiated")
$monitorQueue.updateQueue()
$monitorQueue.start()