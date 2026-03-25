# CertificateOps.psm1
# Certificate operations for code signing (Phase 1: PFX, Phase 2: Windows Trusted Signing)
# Requires PublishCore.psm1 to be imported first for Write-PublishLog and Test-ExecutionContext functions

<#
.SYNOPSIS
    Gets or imports the signing certificate

.DESCRIPTION
    Context-aware certificate loading:
    - Local: Uses certificate from Windows certificate store by thumbprint
    - CI/CD: Imports PFX file from secrets

.PARAMETER Thumbprint
    The certificate thumbprint to locate

.PARAMETER PfxPath
    Path to PFX file (optional, for CI/CD)

.PARAMETER PfxPassword
    SecureString password for PFX file (optional, for CI/CD)

.OUTPUTS
    X509Certificate2 - The certificate object
#>
function Get-SigningCertificate {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Thumbprint,

        [Parameter(Mandatory=$false)]
        [string]$PfxPath = $null,

        [Parameter(Mandatory=$false)]
        [securestring]$PfxPassword = $null
    )

    $context = Test-ExecutionContext

    Write-PublishLog "Getting signing certificate (Context: $context)..." "Info"
    Write-PublishLog "Certificate thumbprint: $Thumbprint" "Info"

    if ($context -eq "GitHubActions") {
        # CI/CD: Import from PFX file
        Write-PublishLog "CI/CD context detected - importing certificate from PFX" "Info"

        if (-not $PfxPath) {
            throw "PFX path is required in CI/CD context"
        }

        if (-not $PfxPassword) {
            throw "PFX password is required in CI/CD context"
        }

        if (-not (Test-Path $PfxPath)) {
            throw "PFX file not found at: $PfxPath"
        }

        # Import PFX to CurrentUser certificate store
        $cert = Import-PfxCertificate -FilePath $PfxPath -Password $PfxPassword -CertStoreLocation "Cert:\CurrentUser\My"

        Write-PublishLog "Certificate imported successfully from PFX" "Success"
        Write-PublishLog "Certificate subject: $($cert.Subject)" "Info"
        Write-PublishLog "Certificate thumbprint: $($cert.Thumbprint)" "Info"
        Write-PublishLog "Certificate expires: $($cert.NotAfter)" "Info"

        return $cert
    } else {
        # Local: Ensure certificate exists in store before attempting MSBuild signing
        Write-PublishLog "Local context detected - checking Windows certificate store" "Info"

        $normalizedThumbprint = ($Thumbprint -replace "\s", "").ToUpperInvariant()

        $certPath = "Cert:\CurrentUser\My\$normalizedThumbprint"
        if (-not (Test-Path $certPath)) {
            # Additional hint: it might be installed under LocalMachine
            $localMachinePath = "Cert:\LocalMachine\My\$normalizedThumbprint"
            $isInLocalMachine = Test-Path $localMachinePath

            Write-PublishLog "Signing certificate not found in CurrentUser\\My store." "Error"
            Write-PublishLog "Expected thumbprint: $normalizedThumbprint" "Error"
            if ($isInLocalMachine) {
                Write-PublishLog "Note: A certificate with this thumbprint WAS found in LocalMachine\\My." "Warning"
                Write-PublishLog "MSBuild/ClickOnce signing often expects the cert in CurrentUser\\My (Personal) for the publishing user." "Warning"
            }

            Write-PublishLog "To fix:" "Info"
            Write-PublishLog "  1) Import the PFX into *Current User* certificate store (Personal):" "Info"
            Write-PublishLog "     - Double-click the .pfx -> Install Certificate... -> Current User -> Place in 'Personal'" "Info"
            Write-PublishLog "     - Or run: Import-PfxCertificate -FilePath <path-to.pfx> -CertStoreLocation Cert:\\CurrentUser\\My" "Info"
            Write-PublishLog "  2) Ensure the imported cert includes the PRIVATE KEY (you should see a key icon in certmgr.msc)." "Info"
            Write-PublishLog "  3) Verify the thumbprint in certmgr.msc (Certificates - Current User -> Personal -> Certificates)." "Info"
            Write-PublishLog "  4) Re-run Publish.ps1" "Info"

            throw "Certificate with thumbprint $normalizedThumbprint is not installed in Cert:\\CurrentUser\\My"
        }

        Write-PublishLog "Certificate found in CurrentUser\\My: $certPath" "Success"
        $cert = Get-Item $certPath

        # Validate private key presence (required for signing)
        if (-not $cert.HasPrivateKey) {
            Write-PublishLog "Certificate found but does NOT have a private key (cannot sign)." "Error"
            Write-PublishLog "Re-import the PFX and ensure 'Mark this key as exportable' is enabled if needed." "Info"
            throw "Certificate $normalizedThumbprint is missing a private key"
        }

        Write-PublishLog "Certificate loaded successfully from Windows store" "Success"
        Write-PublishLog "Certificate subject: $($cert.Subject)" "Info"
        Write-PublishLog "Certificate thumbprint: $($cert.Thumbprint)" "Info"
        Write-PublishLog "Certificate expires: $($cert.NotAfter)" "Info"

        return $cert
    }
}

<#
.SYNOPSIS
    Imports a PFX certificate

.DESCRIPTION
    Imports a PFX file into the specified certificate store

.PARAMETER FilePath
    Path to the PFX file

.PARAMETER Password
    SecureString password for the PFX file

.PARAMETER CertStoreLocation
    Certificate store location (default: Cert:\CurrentUser\My)

.OUTPUTS
    X509Certificate2 - The imported certificate
#>
function Import-PfxCertificate {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,

        [Parameter(Mandatory=$true)]
        [securestring]$Password,

        [Parameter(Mandatory=$false)]
        [string]$CertStoreLocation = "Cert:\CurrentUser\My"
    )

    Write-PublishLog "Importing PFX certificate from: $FilePath" "Info"

    if (-not (Test-Path $FilePath)) {
        throw "PFX file not found: $FilePath"
    }

    try {
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new()
        $cert.Import($FilePath, $Password, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)

        # Add to store
        $store = [System.Security.Cryptography.X509Certificates.X509Store]::new("My", [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $store.Add($cert)
        $store.Close()

        Write-PublishLog "Certificate imported successfully" "Success"
        Write-PublishLog "Thumbprint: $($cert.Thumbprint)" "Info"

        return $cert
    } catch {
        throw "Failed to import PFX certificate: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Validates certificate

.DESCRIPTION
    Checks if certificate is valid and not expired

.PARAMETER Certificate
    The certificate to validate

.OUTPUTS
    Boolean - True if valid
#>
function Test-CertificateValidity {
    param(
        [Parameter(Mandatory=$true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    Write-PublishLog "Validating certificate..." "Info"

    $now = Get-Date

    # Check if certificate has expired
    if ($Certificate.NotAfter -lt $now) {
        Write-PublishLog "Certificate has expired on $($Certificate.NotAfter)" "Error"
        return $false
    }

    # Check if certificate is not yet valid
    if ($Certificate.NotBefore -gt $now) {
        Write-PublishLog "Certificate is not yet valid (valid from $($Certificate.NotBefore))" "Error"
        return $false
    }

    # Warn if certificate is expiring soon (within 30 days)
    $expirationDays = ($Certificate.NotAfter - $now).Days
    if ($expirationDays -lt 30) {
        Write-PublishLog "WARNING: Certificate expires in $expirationDays days!" "Warning"
    }

    Write-PublishLog "Certificate is valid (expires: $($Certificate.NotAfter))" "Success"
    return $true
}

<#
.SYNOPSIS
    Exports certificate to PFX file

.DESCRIPTION
    Exports a certificate from the store to a PFX file (for backup or CI/CD setup)

.PARAMETER Thumbprint
    The certificate thumbprint to export

.PARAMETER OutputPath
    Path where the PFX file will be saved

.PARAMETER Password
    SecureString password for the PFX file
#>
function Export-CertificateToPfx {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Thumbprint,

        [Parameter(Mandatory=$true)]
        [string]$OutputPath,

        [Parameter(Mandatory=$true)]
        [securestring]$Password
    )

    Write-PublishLog "Exporting certificate to PFX..." "Info"

    $certPath = "Cert:\CurrentUser\My\$Thumbprint"
    if (-not (Test-Path $certPath)) {
        throw "Certificate not found in Windows certificate store at: $certPath"
    }

    $cert = Get-Item $certPath

    try {
        Export-PfxCertificate -Cert $cert -FilePath $OutputPath -Password $Password | Out-Null
        Write-PublishLog "Certificate exported successfully to: $OutputPath" "Success"

        # Base64 encode for GitHub secret
        $pfxBytes = [System.IO.File]::ReadAllBytes($OutputPath)
        $base64 = [Convert]::ToBase64String($pfxBytes)

        Write-Host ""
        Write-Host "=== FOR GITHUB SECRET ===" -ForegroundColor Cyan
        Write-Host "Copy the base64 string below and paste it into GitHub secret: CODE_SIGNING_PFX_BASE64" -ForegroundColor Yellow
        Write-Host ""
        Write-Host $base64 -ForegroundColor Green
        Write-Host ""
        Write-Host "=========================" -ForegroundColor Cyan

    } catch {
        throw "Failed to export certificate: $($_.Exception.Message)"
    }
}

# Export functions
Export-ModuleMember -Function @(
    'Get-SigningCertificate',
    'Import-PfxCertificate',
    'Test-CertificateValidity',
    'Export-CertificateToPfx'
)
