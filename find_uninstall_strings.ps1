#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Date: 4/10/2019
# Description: Searches for the 32bit and 64bit uninstall strings of a specified application
#####
param(
    [Parameter(Mandatory=$true)][string] $appName
)

$64bitSoftware = Get-ChildItem -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" | 
    Where-Object {($_.GetValue("displayname")) -and ($_.GetValue("uninstallstring"))}
$64bitSoftware = $64bitSoftware | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString |
    Where-Object {$_.DisplayName -match $appName}

$32bitSoftware = Get-ChildItem -Path "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | 
    Where-Object {($_.GetValue("displayname")) -and ($_.GetValue("uninstallstring"))}
$32bitSoftware = $32bitSoftware | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString |
    Where-Object {$_.DisplayName -match $appName}

If ($64bitSoftware -or $32bitSoftware) {
    If ($64bitSoftware) {
        Write-Host "---------------"
        Write-Host "64-Bit Software"
        Write-Host "---------------"

        foreach ($app in $64bitSoftware) {
            Write-Host " - $($app.DisplayName): $($app.UninstallString)"
        }

        Write-Host ""
    }

    If ($32bitSoftware) {
        Write-Host "---------------"
        Write-Host "32-Bit Software"
        Write-Host "---------------"

        foreach ($app in $32bitSoftware) {
            Write-Host " - $($app.DisplayName): $($app.UninstallString)"
        }
    }

} else {
    Write-Host "No results found for the application: $appName"
}