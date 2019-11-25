#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Date: 11/12/2019
# Description: Queries specified OUs within AD to check for duplicates entries
# based upon service tag (extensionAttribute4) as well as check for computer names
# not aligned with naming scheme. Once finished, output the results and export to
# a csv.
#####
Import-Module ActiveDirectory

# Array of OUs to check for duplicates and naming inconsistancies
$OUsToCheck = @(
    "OU=Office,OU=UWEC Computers,DC=UWEC, DC=edu",
    "OU=Labs,OU=UWEC Computers,DC=UWEC, DC=edu",
    "OU=Embedded,OU=UWEC Computers,DC=UWEC, DC=edu")

$duplicates = New-Object System.Collections.ArrayList
$inconsistent = New-Object System.Collections.ArrayList

# Get all computers in each OU and check for duplicates and naming inconsistancies
Foreach ($OU in $OUsToCheck) {
    $computers = Get-ADComputer -Filter * -SearchBase $OU -Properties Name, DistinguishedName, pwdLastSet, Enabled, extensionAttribute1, extensionAttribute4
    $computersToCheck = [System.Collections.ArrayList] $computers.Clone() # Clone so we can manipulate array while iterating

    Foreach ($computer in $computers) {
        # Remove the current computer from the list before passing it along to reduce workload
        $computersToCheck.Remove($computer)

        # Check if duplicate and that service tag isn't empty or N/A
        $duplicateComputers = $($computersToCheck | Where-Object { ($_.extensionAttribute4 -eq $computer.extensionAttribute4) -and !($([string]::IsNullOrWhiteSpace($computer.extensionAttribute4)) -or $($computer.extensionAttribute4 -eq "N/A")) })
        If ($($duplicateComputers | Measure-Object).Count -gt 0) {
            $computer.Duplicates = $duplicateComputers | Foreach-Object { "$($_.Name)" } # Create a string containing all duplicate computer names
            $duplicates.Add($computer) > $null # Suppress console output when appending
        }

        # Check if name is inconsistent with the exception of Lab computers
        If ($($OU -ne "OU=Labs,OU=UWEC Computers,DC=UWEC, DC=edu") -and !($computer.Name -match $computer.extensionAttribute1)) {
            $inconsistent.Add($computer) > $null # Suppress console output when appending
        }
    }
}

# Display results
Write-Host "<-- Duplicate Service Tag -->"
$duplicates | Format-Table -Property Name, DistinguishedName, @{Name = "pwdLastSet"; Expression = { ([datetime]::FromFileTime($_.pwdLastSet)) } }, Enabled, @{Name = "Asset Tag"; Expression = { $_.extensionAttribute1 } }, @{Name = "Service Tag"; Expression = { $_.extensionAttribute4 } }, @{Name = "Duplicates"; Expression = { $_.Duplicates -join ", " } } | Out-String | Write-Host

Write-Host "<-- Inconsistent PC Names -->"
$inconsistent | Format-Table -Property Name, DistinguishedName, @{Name = "pwdLastSet"; Expression = { ([datetime]::FromFileTime($_.pwdLastSet)) } }, Enabled, @{Name = "Asset Tag"; Expression = { $_.extensionAttribute1 } }, @{Name = "Service Tag"; Expression = { $_.extensionAttribute4 } } | Out-String | Write-Host

# Export results as CSV
$duplicates | Select-Object -Property Name, DistinguishedName, @{Name = "pwdLastSet"; Expression = { ([datetime]::FromFileTime($_.pwdLastSet)) } }, Enabled, @{Name = "Asset Tag"; Expression = { $_.extensionAttribute1 } }, @{Name = "Service Tag"; Expression = { $_.extensionAttribute4 } }, @{Name = "Duplicates"; Expression = { $_.Duplicates -join ", " } } | Export-Csv -Path ./duplicates.csv -NoTypeInformation
$inconsistent | Select-Object -Property Name, DistinguishedName, @{Name = "pwdLastSet"; Expression = { ([datetime]::FromFileTime($_.pwdLastSet)) } }, Enabled, @{Name = "Asset Tag"; Expression = { $_.extensionAttribute1 } }, @{Name = "Service Tag"; Expression = { $_.extensionAttribute4 } } | Export-Csv -Path ./inconsistent.csv -NoTypeInformation