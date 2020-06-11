#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Date: 06/11/2020
# Description: Searched through all Task Sequences for steps matching the provided name,
# prompting for deletion upon discovery.
#####
param(
    [string] $stepToBeRemoved
)

If ( !$stepToBeRemoved ) {
    Write-Host -F Red "ERROR: Please provided a partial string to look for."
    Write-Host 'Usage: taskseq_remove_step.ps1 "step name"'
    Exit 0
}

Try {
    Write-Host "Loading Configuration Manager Module..."
    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
} Catch {
    Write-Host "Configuration Manager Module failed to load. Exiting..."
    Exit 1
}

$currentLocation = Get-Location
$siteCode = Get-PSDrive -PSProvider CMSITE
Set-Location -Path "$($SiteCode.Name):\"

Write-Host "Retrieving all task sequences. This may take a moment..."
$tasks = Get-CMTaskSequence | Select-Object Name, PackageID

Write-host "Tasks retrieved. Checking task sequences for matches..."
ForEach ($task in $tasks) {
    Write-Host "Checking $($task.Name)..."
    $steps = Get-CMTaskSequenceStep -TaskSequenceID $task.PackageID

    Foreach ($step in $steps) {
        If ($step.Name -match $stepToBeRemoved) {
            Write-Host "<==========>"
            Write-Host -F Green "Match Found: $($task.Name) > $($step.Name)"
            
            $showDetails = Read-Host -Prompt "Show step details? [Y/N]"
            If ($showDetails -match "y" -or $showDetails -match "yes") {
                Write-Host "Details:"
                $step | Format-List
            }

            Write-Host "Confirm deletion..."
            Remove-CMTaskSequenceStep -TaskSequenceID $task.PackageID -StepName $step.Name
            Write-Host "<==========>"
        }
    }
}

Set-Location $currentLocation