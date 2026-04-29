function Get-ExpiringCertificateReport {
    <#
        .SYNOPSIS
        Generate a report of expiring certificates from an Active Directory Certificate Services Certificate Authority.

        .DESCRIPTION
        This script checks ADCS Certificate Authorities for issued certificate requests that are expiring within the configured lead time.
        Specify a list of certificate templates to include, and it will resolve the provided identifiers (for example OID, short
        name, or DisplayName), find expiring certs using those templates, and then send a report as directed.

        .PARAMETER Recipients
        To-Do: Add a function parameter to send an email to specific recipients.

        .PARAMETER Output
        To-Do: Add a function parameter to choose email, HTML, CSV, JSON, or console output.

        .INPUTS
        None. You cannot pipe objects to this script.

        .OUTPUTS
        Email
        HTML, CSV, JSON, or XML file
        Console

        .NOTES
        Author:     Sam Erde
        Modified:   2023/07/21

        Depends on the PSPKI module at https://www.powershellgallery.com/packages/PSPKI and the AD Certificate Services RSAT feature.

        To Do: Add checks for prerequisites, turn into function(s), take parameters for recipients and report output type,
        get CAs in all domains in AD forest, error handling, show all template names (and optionally use Out-GridView/Out-
        ConsoleGridView to select desired templates), use OGV to generate a text file containing templates and then use
        that file as list of monitored certificate templates for expiring certificates report.
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost')]
    param ()

    $Version = '2023.07.21'
    $Header = @"

███╗   ██╗ ██████╗      ██████╗███████╗██████╗ ████████╗
████╗  ██║██╔═══██╗    ██╔════╝██╔════╝██╔══██╗╚══██╔══╝
██╔██╗ ██║██║   ██║    ██║     █████╗  ██████╔╝   ██║
██║╚██╗██║██║   ██║    ██║     ██╔══╝  ██╔══██╗   ██║
██║ ╚████║╚██████╔╝    ╚██████╗███████╗██║  ██║   ██║
╚═╝  ╚═══╝ ╚═════╝      ╚═════╝╚══════╝╚═╝  ╚═╝   ╚═╝

██╗     ███████╗███████╗████████╗    ██████╗ ███████╗██╗  ██╗██╗███╗   ██╗██████╗
██║     ██╔════╝██╔════╝╚══██╔══╝    ██╔══██╗██╔════╝██║  ██║██║████╗  ██║██╔══██╗
██║     █████╗  █████╗     ██║       ██████╔╝█████╗  ███████║██║██╔██╗ ██║██║  ██║
██║     ██╔══╝  ██╔══╝     ██║       ██╔══██╗██╔══╝  ██╔══██║██║██║╚██╗██║██║  ██║
███████╗███████╗██║        ██║       ██████╔╝███████╗██║  ██║██║██║ ╚████║██████╔╝
╚══════╝╚══════╝╚═╝        ╚═╝       ╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝╚═╝  ╚═══╝╚═════╝

v$Version

"@
    Write-Host -ForegroundColor Cyan -BackgroundColor Black $Header

    # ════════════════════════════════════════════════════════╗
    # Modify these variables to suit your environment:        ║
    #   To, From, SMTPServer, DaysLeft, TemplateNamesIncluded ║

    $To = @('recipient1@example.com', 'recipient2@example.com')
    $From = 'NoCertLeftBehind@example.com'
    $SMTPServer = 'smtp.example.com'

    # Change "DaysLeft" to whatever lead time you want for the notification of expiring certificates.
    $DaysLeft = 30

    # List the display names of the certificate templates that you want to monitor using a multi-line here-string that is converted to an array.
    # This is simply easier than typing every name in quotes and separating them with a comma.
    $TemplateNamesIncluded = (@'
Root Certification Authority
CA Exchange
CEP Encryption
Code Signing
Cross Certification Authority
Trust List Signing
Directory Email Replication
Domain Controller
Domain Controller Authentication
Enrollment Agent
Exchange Enrollment Agent (Offline request)
Exchange Signature Only
IPSec (Offline request)
IPSec
Kerberos Authentication
Key Recovery Agent
Enrollment Agent (Computer)
RAS and IAS Server
SCEP
Signature with Key Encipherment
Subordinate Certification Authority
Web Server
WSUS Signing Certificate
'@).Split([Environment]::NewLine) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    # End of Customizations ║
    # ══════════════════════╝

    # Import required modules and Windows features
    if (Get-Module -Name 'PSPKI' -ListAvailable) {
        Write-Information 'The PSPKI module is installed.'
    } else {
        Write-Information 'The PSPKI module is not installed. Attempting installation...'
        try {
            Install-Module -Name PSPKI -AllowClobber -Scope CurrentUser -Force
        } catch {
            Write-Error 'PSPKI module installation failed.'
        }
    }
    if ( (Get-WindowsCapability -Online -Name 'Rsat.CertificateServices.Tools~~~~0.0.1.0').State -eq 'Installed') {
        Write-Information 'The Certificate Services RSAT feature is installed.'
    } else {
        try {
            Get-WindowsCapability -Online -Name 'Rsat.CertificateServices.Tools~~~~0.0.1.0' | Add-WindowsCapability -Online
        } catch {
            Write-Error 'Failed to install the Certificate Services RSAT feature. Please do so manually.'
        }
    }
    # End of module installation check

    # ═══════════════════════════════════════════════════════════════════════════════════════╗
    # Shortcut (code snippet) variables: these make the following code simpler to work with: ║

    # Extract just the certificate authority's hostname from its configuration name.  Used within Get-CertificateRequests below.
    $CaName = @{ Name = 'CA'; Expression = { $_.ConfigString.Split('\')[1] } }

    # Certificate age filter statement.  Used within Get-CertificateRequests below.
    $CertAgeFilter = "NotAfter -ge $(Get-Date)", "NotAfter -le $((Get-Date).AddDays($DaysLeft))"

    $TemplateDisplayNameByIdentifier = @{}

    # Translate a certificate template identifier to its readable display name. Used within Get-CertificateRequests below.
    $CertTemplateName = @{ Name = 'TemplateName'; Expression = {
            $TemplateIdentifier = $_.CertificateTemplate
            if (-not [string]::IsNullOrWhiteSpace($TemplateIdentifier) -and $TemplateDisplayNameByIdentifier.ContainsKey($TemplateIdentifier)) {
                $TemplateDisplayNameByIdentifier[$TemplateIdentifier]
            } elseif ($TemplateIdentifier -like '1*') {
                $CertificateTemplate = Get-CertificateTemplate -OID $TemplateIdentifier
                if ($null -ne $CertificateTemplate) {
                    $CertificateTemplate.DisplayName
                } else {
                    $TemplateIdentifier
                }
            } else {
                $TemplateIdentifier
            }
        }
    } # End CertTemplateName

    $Domain = $([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)

    # End of shortcut (code snippet) variables ║
    # ═════════════════════════════════════════╝


    # ══════════════════╗
    # Collect the data: ║

    # Find certificate authorities in the domain and get their hostname[s]
    Write-Information "Finding certificate authorities in $Domain..."
    $CANames = (Get-CA | Select-Object Computername).Computername
    if ($CANames.Count -lt 1) {
        Write-Warning 'No certificate authorities were found in Active Directory.'
        return
    }
    Write-Information "Found: $CANames `n"

    # Get all identifiers for the certificate templates that you want to monitor.
    # Some issued requests expose the template short name, such as SubCA, instead of an OID.
    Write-Information "Getting identifiers for $($TemplateNamesIncluded.Count) certificate templates..."
    $TemplateIdentifiersIncluded = @(
        foreach ($Item in $TemplateNamesIncluded) {
            try {
                $CertificateTemplate = Get-CertificateTemplate -DisplayName $Item
                if ($null -eq $CertificateTemplate) {
                    Write-Warning "Certificate template '$Item' was not found or could not be retrieved."
                    continue
                }

                $TemplateIdentifiers = @(
                    $CertificateTemplate.Oid.Value
                    $CertificateTemplate.Name
                    $CertificateTemplate.DisplayName
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

                foreach ($TemplateIdentifier in $TemplateIdentifiers) {
                    $TemplateDisplayNameByIdentifier[$TemplateIdentifier] = $CertificateTemplate.DisplayName
                }

                $TemplateIdentifiers
            } catch {
                Write-Warning "Unable to resolve certificate template '$Item'. Skipping. Error: $_"
                continue
            }
        }
    ) | Select-Object -Unique

    if ($TemplateIdentifiersIncluded.Count -eq 0) {
        Write-Warning 'No certificate template identifiers were found.'
        return
    }

    # Get all relevant certificates that are expiring within the configured lead time.
    Write-Information "`nGetting certificates that are expiring in the next $DaysLeft days from $CANames..."
    $Certificates = ( Get-IssuedRequest -CertificationAuthority $CANames -Property [Request.RequesterName], CertificateHash -Filter $CertAgeFilter ).Where(
        { $TemplateIdentifiersIncluded -contains $_.CertificateTemplate } ) |
        Select-Object *, $CaName, $CertTemplateName, @{Name = 'Thumbprint'; Expression = { $_.CertificateHash -Replace (' ', '') } }
    # In the above line, $CaName and $CertTemplateName are "shortcut snippet variables" like a function that formats or translates the desired output.

    if ($Certificates.Count -eq 0) {
        Write-Information 'No certificates were found to report on.'
        return
    } else {
        Write-Information "Found $($Certificates.Count) certificates that expire within $DaysLeft days..."
    }

    # Done getting the certificates. ║
    # ═══════════════════════════════╝


    # ═══════════════════════════╗
    # Build and send the report: ║

    # Create a structured table with certificate details. Designed specifically for an HTML-based email.
    $Table = [System.Data.DataTable]::New('CertificatesTable')
    @(
        'Name'
        'Expiration'
        'Identifiers'
        'Requester'
        'CA'
    ) | ForEach-Object { $Table.Columns.Add($_) | Out-Null }

    # Add each certificate to the table
    foreach ($Certificate in $Certificates) {
        $CertificateRow = $Table.NewRow()
        $CertificateRow.Name = "$($Certificate.CommonName)╗($($Certificate.Templatename))"
        $CertificateRow.Expiration = $Certificate.NotAfter
        $CertificateRow.Requester = $Certificate.'Request.RequesterName'
        $CertificateRow.Identifiers = "Serial: $($Certificate.SerialNumber)╗Thumbprint: $($Certificate.Thumbprint)╗Request ID: $($Certificate.RequestID)"
        $CertificateRow.CA = $Certificate.CA
        $Table.Rows.Add($CertificateRow)
    }

    $HtmlHeader = @'
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TD {border-width: 1px; padding: 4px; border-style: solid; border-color: black;}
</style>
'@
    $PreContent = "Internally-issued certificates that will expire in the next $DaysLeft days: ╗╗"
    $EmailHtml = ($Table | ConvertTo-Html -Head $HtmlHeader -PreContent $PreContent -Property Name, Expiration, Identifiers, Requester, CA | Out-String).Replace('╗', '<br/>')
    $Subject = "$Domain Certificate Expiration Report"
    Send-MailMessage -To $To -From $From -SmtpServer $SMTPServer -Subject $Subject -Body $EmailHtml -BodyAsHtml
}
