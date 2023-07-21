<#
    .SYNOPSIS
    Generate a report of exiring certificates from an Active Directory Certificate Services Certificate Authority.

    .DESCRIPTION
    This script checks ADCS Certificate Authorities for issued certificate requests that are expiring in the next 45 days.
    Specify a list of template names to include, and it will translate that to their OIDs, find expiring certs using those
    templates, and then send a report as directed.

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
$Version = "2023.07.21"
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
Write-Host -ForegroundColor Cyan $Header

# ════════════════════════════════════════════════════════╗
# Modify these variables to suit your environment:        ║
#   To, From, SMTPServer, DaysLeft, TemplateNamesIncluded ║

$To = @('recipient1@example.com','recipient2@example.com')
$From = 'NoCertLeftBehind@example.com'
$SMTPServer = 'smtp.example.com'

# Change "DaysLeft" to whatever lead time you want for the notification of expiring certificates.
$DaysLeft = 45

# List the display names of the certificate templates that you want to monitor using a multi-line here-string that is converted to an array.
# This is simply easier than typing every name in quotes and separating them with a comma.
$TemplateNamesIncluded = (@"
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
"@).Split([Environment]::NewLine)

# End of Customizations ║
# ══════════════════════╝

Import-Module PSPKI

# ═══════════════════════════════════════════════════════════════════════════════════════╗
# Shortcut (code snippet) variables: these make the following code simpler to work with: ║

# Extract just the certificate authority's hostname from its configuration name.  Used within Get-CertificateRequests below.
$CaName = @{ Name="CA"; Expression = {$_.ConfigString.Split("\")[1]} }

# Certificate age filter statement.  Used within Get-CertificateRequests below.
$CertAgeFilter = "NotAfter -ge $(Get-Date)", "NotAfter -le $((Get-Date).AddDays($DaysLeft))"

# Translate a certificate template OID to its readable display name. Used within Get-CertificateRequests below.
$CertTemplateName = @{ Name="TemplateName"; Expression = {
        if ($_.CertificateTemplate -like "1*") {
            (Get-CertificateTemplate -OID $_.CertificateTemplate).DisplayName
        }
        else {
            $_.CertificateTemplate
        }
    }
} # End CertTemplateName

$Domain = $([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)

# End of shortcut (code snippet) variables ║
# ═════════════════════════════════════════╝


# ══════════════════╗
# Collect the data: ║

# Find certificate authorities in the domain and get their hostname[s]
Write-Host "Finding certificate authorities in $Domain..."
$CANames = (Get-CA | Select-Object Computername).Computername
if ($CANames.Count -lt 1) {
    Write-Warning "No certificate authorities were found in Active Directory."
    Break
}
Write-Host "Found: $CANames `n"

# Get OIDs of all certificate templates that you want to monitor. The script takes MUCH longer when querying by template name.
Write-Host "Getting OIDs for $($TemplateNamesIncluded.Count) certificate templates..."
$TemplateOidsIncluded = foreach ($item in $TemplateNamesIncluded) {
    try {
        (Get-CertificateTemplate -DisplayName $item).Oid.Value
    }
    catch {
        Write-Host "Unable to get the certificate templates. Please review the error and try again."
        Write-Error $error
        Break
    }
}

# Get all relevant certificates that are expiring within the next 45 days.
Write-Host `n"Getting certifictes that are expiring in the next $DaysLeft days from $CANames..."
$Certificates = ( Get-IssuedRequest -CertificationAuthority $CANames -Property [Request.RequesterName], CertificateHash -Filter $CertAgeFilter ).Where( 
    { $TemplateOidsIncluded -imatch $_.CertificateTemplate } ) | 
    Select-Object *, $CaName, $CertTemplateName, @{Name="Thumbprint"; Expression = { $_.CertificateHash -Replace (' ','') } }
    # In the above line, $CaName and $CertTemplateName are "shortcut snippet variables" like a function that formats or translates the desired output.

if ($certificates.Length -eq 0) { 
    Write-Output "No certificates were found to report on."
    Break
}
else {
    Write-Host -ForegroundColor Yellow "Found $($Certificates.Count) certificates that expire within $DaysLeft days..."
}

# Done getting the certificates. ║
# ═══════════════════════════════╝


# ═══════════════════════════╗
# Build and send the report: ║

# Create a structured table with certificate details. Designed specifically for an HTML-based email.
$table = [System.Data.DataTable]::New("CertificatesTable")
@(
  "Name"
  "Expiration"
  "Identifiers"
  "Requester"
  "CA"
) | ForEach-Object { $table.Columns.Add($_) | Out-Null }

# Add each certificate to the table
foreach ($cert in $certificates) {
    $certRow = $table.NewRow()
        $certRow.Name           = "$($cert.CommonName)╗($($cert.Templatename))"
        $certRow.Expiration     = $cert.NotAfter
        $certRow.Requester      = $cert."Request.RequesterName"
        $certRow.Identifiers    = "Serial: $($cert.SerialNumber)╗Thumbprint: $($cert.Thumbprint)╗Request ID: $($cert.RequestID)"
        $certRow.CA             = $cert.CA
    $table.Rows.Add($CertRow)
}

$HtmlHeader = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TD {border-width: 1px; padding: 4px; border-style: solid; border-color: black;}
</style>
"@
$PreContent = "Internally-issued certificates that will expire in the next $DaysLeft days: ╗╗"
$EmailHtml = ($table | ConvertTo-Html -Head $HtmlHeader -PreContent $PreContent -Property Name,Expiration,Identifiers,Requester,CA | Out-String).Replace('╗','<br/>')
$Subject = "$Domain Certificate Expiration Report"
Send-MailMessage -To $To -From $From -SmtpServer $SMTPServer -Subject $Subject -Body $EmailHtml -BodyAsHTML
