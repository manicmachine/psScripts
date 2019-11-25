#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Date: 9/11/2019
# Description: Queries SCCM for members of a specified device collection and 
# then passes those results to .\login_status.ps1 to begin monitoring for availability
#####
param(
    [string] $collection
)

Import-Module "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
$currentLocation = Get-Location
Set-Location -Path "WEC:"

$members = (Get-CMCollectionMember -CollectionName $collection | Select-Object -ExpandProperty Name) -join ", "
Set-Location -Path $currentLocation

Invoke-Expression -Command "$PSScriptRoot\login_status.ps1 -Computers $members"