# No Cert Left Behind
Generate a report of exiring certificates from your Active Directory Certificate Services Certificate Authority.

This script checks ADCS Certificate Authorities for issued certificate requests that are expiring in the next 45 days.

Specify a list of template names to include, and it will translate that to their OIDs, find expiring certs using those templates, and then send a report as directed.

Depends on the PSPKI module at https://www.powershellgallery.com/packages/PSPKI and the AD Certificate Services RSAT feature.

To Do: 
- [ ] Add checks for prerequisites
- [ ] Turn into function(s)
- [ ] Take parameters for recipients and report output type
- [ ] Get CAs in all domains in AD forest
- [ ] Add error handling
- [ ] Sshow all template names (and optionally use Out-GridView/Out-ConsoleGridView to select desired templates)
- [ ] Use OGV to generate a text file containing templates and then use that file as list of monitored certificate templates for expiring certificates report
