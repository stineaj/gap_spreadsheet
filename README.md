# gap_spreadsheet
python script to generate a formatted excel spreadsheet for dbgap uploads

Pulls the samples from dbGaP Sample Status CGI and outputs a spreadsheet in the CWD that includes the dbGaP study accession and sample info necessary to submit data to SRA for the project.

Currently only outputs the samples that have a status of 'Loaded'.  

Possible Additions
- Quiet/Verbose mode
- Page for already loaded data statistics
- Updated instructions if the spreadsheet is more integrated with pipeline
