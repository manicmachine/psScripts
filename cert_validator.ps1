#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Date: 3/26/2020
# Description: Checks to see if there is a machine certificate present issued to the
# correct computer name. If so, check if it's still valid. If no certificates are
# present or none are valid, request a new machine certificate and display the 
# result along with certificate details.
#####
# Get list of installed local computer certificates issued to the current computer name
$localCerts = $(Get-ChildItem -Path "Cert:\LocalMachine\My") -match $env:COMPUTERNAME

# Check if any are still valid
$notExpired = $false
$validCerts = [System.Collections.ArrayList]::new()
$date = Get-Date

ForEach ($cert in $localCerts) {
    If ($cert.NotAfter -gt $date) {
        $notExpired = $true
        $validCerts.Add($cert) | Out-Null
    }
}

# Request a new machine certificate if none are issued to the correct computer name or if expired
If (($localCerts.Count -lt 1) -or !$notExpired) {
    Write-Host "No valid machine certificates present. Requesting new certifcate..."
    $result = Get-Certificate -Template "Machine" -Url "ldap:" -CertStoreLocation "Cert:\LocalMachine\My"

    If ($result.Status -eq "Issued") {
        Write-Host "Request successful!" 
        Write-Host "Subject: $($result.Certificate.Subject)"
        Write-Host "Issuer: $($result.Certificate.Issuer)"
        Write-Host "Serial: $($result.Certificate.SerialNumber)"
        Write-Host "Thumbprint: $($result.Certificate.Thumbprint)"
        Write-Host "Valid Until: $($result.Certificate.NotAfter)"
    } Else {
        Write-Host -ForegroundColor Red "Request failed!"
        Write-Host "Status: $($request.Status)"
        Exit 1
    }
} Else {
    # Else, display the valid certificates located
    Write-Host "$($validCerts.Count) valid machine certificate(s) present."
    Foreach ($cert in $validCerts) {
        Write-Host "Subject: $($cert.Subject)"
        Write-Host "Issuer: $($cert.Issuer)"
        Write-Host "Serial: $($cert.SerialNumber)"
        Write-Host "Thumbprint: $($cert.Thumbprint)"
        Write-Host "Valid Until: $($cert.NotAfter)"
    }
}