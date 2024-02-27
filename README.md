# No Cert Left Behind

![Banner image: An old attic cluttered with a wooden chest and shelves full of old papers.](NoCertLeftBehind.jpg)
Image Credit: [Peter Herrmann](https://unsplash.com/@tama66) via Unsplash.

Generate a report of expiring certificates from your Active Directory Certificate Services Certificate Authority.

[![Codacy Badge](https://api.codacy.com/project/badge/Grade/92b18813ce1a4668b804be370fa004bb)](https://app.codacy.com/gh/SamErde/No-Cert-Left-Behind?utm_source=github.com&utm_medium=referral&utm_content=SamErde/No-Cert-Left-Behind&utm_campaign=Badge_Grade)

This script checks ADCS Certificate Authorities for issued certificate requests that are expiring in the next 45 days. Specify a list of the template names that you want to check, and it will translate that to their OIDs, find expiring certs using those templates, and then send a report as directed. It is recommended to ignore certain templates that are always automatically renewed by computer and users.

Depends on the PSPKI module at https://www.powershellgallery.com/packages/PSPKI and the AD Certificate Services RSAT feature.

To Do:

- [ ] Add checks for prerequisites
- [ ] Turn into function(s)
- [ ] Take parameters for recipients and report output type
- [ ] Get CAs in all domains in AD forest
- [ ] Add error handling
- [ ] Show all template names (and optionally use Out-GridView/Out-ConsoleGridView to select desired templates)
- [ ] Use OGV to generate a text file containing templates and then use that file as list of monitored certificate templates for expiring certificates report
