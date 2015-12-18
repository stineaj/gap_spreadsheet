#! /usr/bin/env python

### Import ###
import xlsxwriter
import urllib2
import xml.etree.cElementTree as ET
import argparse


### Input Args ###
parser = argparse.ArgumentParser(description="Get the arguments commandline")
parser.add_argument("-p", "--phs", dest="phs", required=True,
                     help="Specifies the phs namespace for the submission.")
parser.add_argument("-v", "--verbose", dest="verbose", required=False,
                     help="Runs the program in verbose mode.")                     
args = parser.parse_args()
phs = args.phs


### Get Sample Info from dbGap ###
sampinfo = urllib2.urlopen('http://www.ncbi.nlm.nih.gov/projects/gap/cgi-bin/GetSampleStatus.cgi?study_id=' + phs +'&rettype=xml')
tree = ET.parse(sampinfo)
for node in tree.iter('Sample'):
   ssid = node.attrib.get('submitted_sample_id')
   status = node.attrib.get('dbgap_status')


### Create a New Excel file. ###
workbook = xlsxwriter.Workbook(str(phs) + '_submission.xlsx', {'in_memory' : True})


### Set Formatting Types ###
bold = workbook.add_format({'bold' : True})
contact = workbook.add_format({'bold' : True, 'bg_color' : '#4F81BD', 'font_color' : 'white', 'align': 'right'})
required = workbook.add_format({'bold' : True, 'bg_color' : '#4F81BD', 'font_color' : 'white', 'text_wrap': True, 'right' : 1})
aligned = workbook.add_format({'bold' : True, 'bg_color' : 'green', 'font_color' : 'white', 'text_wrap': True, 'right' : 1})
paired = workbook.add_format({'bold' : True, 'bg_color' : '#808080', 'font_color' : 'white', 'text_wrap': True, 'right' : 1})
lists = workbook.add_format({'bold' : True, 'bg_color' : '#FFEB9C', 'font_color': '#9C6500', 'right' : 1}) 
url_format = workbook.add_format ({'font_color': 'blue', 'underline': True})


### Dictionaries and Terms ###
STRATEGY = [
['WGA', 'Random sequencing of the whole genome following non-pcr amplification'],
['WGS', 'Random sequencing of the whole genome'],
['WXS', 'Random sequencing of exonic regions selected from the genome'],
['RNA-Seq', 'Random sequencing of whole transcriptome'],
['ncRNA-Seq', 'Capture of other non-coding RNA types, including post-translation modification types such as snRNA (small nuclear RNA) or snoRNA (small nucleolar RNA), or expression regulation types such as siRNA (small interfering RNA) or piRNA/piwi/RNA (piwi-interacting RNA).'],
['miRNA-Seq', 'Random sequencing of small miRNAs'],
['WCS', 'Random sequencing of a whole chromosome or other replicon isolated from a genome'],
['CLONE', 'Genomic clone based (hierarchical) sequencing'],
['POOLCLONE', 'Shotgun of pooled clones (usually BACs and Fosmids)'],
['AMPLICON', 'Sequencing of overlapping or distinct PCR or RT-PCR products'],
['CLONEEND', 'Clone end (5\', 3\', or both) sequencing'],
['FINISHING', 'Sequencing intended to finish (close) gaps in existing coverage'],
['ChIP-Seq', 'Direct sequencing of chromatin immunoprecipitates'],
['MNase-Seq', 'Direct sequencing following MNase digestion'],
['DNase-Hypersensitivity', 'Sequencing of hypersensitive sites, or segments of open chromatin that are more readily cleaved by DNaseI'],
['Bisulfite-Seq', 'Sequencing following treatment of DNA with bisulfite to convert cytosine residues to uracil depending on methylation status'],
['Tn-Seq', 'Sequencing from transposon insertion sites'],
['EST', 'Single pass sequencing of cDNA templates'],
['FL-cDNA', 'Full-length sequencing of cDNA templates'],
['CTS', 'Concatenated Tag Sequencing'],
['MRE-Seq', 'Methylation-Sensitive Restriction Enzyme Sequencing strategy'],
['MeDIP-Seq', 'Methylated DNA Immunoprecipitation Sequencing strategy'],
['MBD-Seq', 'Direct sequencing of methylated fractions sequencing strategy'],
['FAIRE-seq', 'Formaldehyde Assisted Isolation of Regulatory Elements'],
['SELEX', 'Systematic Evolution of Ligands by EXponential enrichment'],
['RIP-Seq', 'Direct sequencing of RNA immunoprecipitates (includes CLIP-Seq, HITS-CLIP and PAR-CLIP). '],
['ChIA-PET', 'Direct sequencing of proximity-ligated chromatin immunoprecipitates.'],
['OTHER', 'Library strategy not listed (please include additional info in the \"design description\")'],
]

SOURCE = [
['GENOMIC', 'Genomic DNA (includes PCR products from genomic DNA)'],
['TRANSCRIPTOMIC', 'Transcription products or non genomic DNA (EST, cDNA, RT-PCR, screened libraries)'],
['METAGENOMIC', 'Mixed material from metagenome'],
['METATRANSCRIPTOMIC', 'Transcription products from community targets'],
['SYNTHETIC', 'Synthetic DNA'],
['VIRAL RNA', 'Viral RNA'],
['OTHER', 'Library strategy not listed (please include additional info in the \"design description\")'],
]

SELECTION = [
['RANDOM', 'Random selection by shearing or other method'],
['PCR', 'Source material was selected by designed primers'],
['RANDOM PCR', 'Source material was selected by randomly generated primers'],
['RT-PCR', 'Source material was selected by reverse transcription PCR'],
['HMPR', 'Hypo-methylated partial restriction digest'],
['MF', 'Methyl Filtrated'],
['CF-S', 'Cot-filtered single/low-copy genomic DNA'],
['CF-M', 'Cot-filtered moderately repetitive genomic DNA'],
['CF-H', 'Cot-filtered highly repetitive genomic DNA'],
['CF-T', 'Cot-filtered theoretical single-copy genomic DNA'],
['MDA', 'Multiple displacement amplification'],
['MSLL', 'Methylation Spanning Linking Library'],
['cDNA', 'complementary DNA'],
['ChIP', 'Chromatin immunoprecipitation'],
['MNase', 'Micrococcal Nuclease (MNase) digestion'],
['DNAse', 'Deoxyribonuclease (MNase) digestion'],
['Hybrid Selection', 'Selection by hybridization in array or solution'],
['Reduced Representation', 'Reproducible genomic subsets, often generated by restriction fragment size selection, containing a manageable number of loci to facilitate re-sampling'],
['Restriction Digest', 'DNA fractionation using restriction enzymes'],
['5-methylcytidine antibody', 'Selection of methylated DNA fragments using an antibody raised against 5-methylcytosine or 5-methylcytidine (m5C)'],
['MBD2 protein methyl-CpG binding domain', 'Enrichment by methyl-CpG binding domain'],
['CAGE', 'Cap-analysis gene expression'],
['RACE', 'Rapid Amplification of cDNA Ends'],
['size fractionation', 'Physical selection of size appropriate targets'],
['Padlock probes capture method', 'Circularized oligonucleotide probes'],
['other', 'Other library enrichment, screening, or selection process (please include additional info in the \"design description\")'],
['unspecified', 'Library enrichment, screening, or selection is not specified (please include additional info in the \"design description\")'],
]

PLATFORMS = {
"_LS454" : ( "454 GS", "454 GS 20", "454 GS FLX", "454 GS FLX+", "454 GS FLX Titanium", "454 GS Junior", "unspecified" ),
"ILLUMINA" : ( "Illumina Genome Analyzer", "Illumina Genome Analyzer II", "Illumina Genome Analyzer IIx", 
"Illumina HiScanSQ", "Illumina HiSeq 2500", "Illumina HiSeq 2000", "Illumina HiSeq 1500", "Illumina HiSeq 1000", "Illumina MiSeq", "Illumina NextSeq 550","Illumina NextSeq 500", "HiSeq X Ten", "unspecified" ),
"HELICOS" : ( "Helicos HeliScope", "unspecified" ),
"ABI_SOLID" : ( "AB SOLiD System", "AB SOLiD System 2.0", "AB SOLiD System 3.0", "AB SOLiD 4 System", "AB SOLiD 4hq System", "AB SOLiD 3 Plus System",
"AB SOLiD PI System", "AB 5500 Genetic Analyzer", "AB 5500xl Genetic Analyzer", "AB 5500xl-W Genetic Analyzer", "unspecified" ),
"COMPLETE_GENOMICS" : ( "Complete Genomics" ),
"OXFORD_NANOPORE" : ( "MinION", "GridION", "unspecified" ),
"PACBIO_SMRT" : ( "PacBio RS", "PacBio RS II", "unspecified" ),
"ION_TORRENT" : ( "Ion Torrent PGM", "Ion Torrent Proton", "unspecified" ),
"CAPILLARY" : ( "AB 3730xL Genetic Analyzer", "AB 3730 Genetic Analyzer", "AB 3500xL Genetic Analyzer",
"AB 3500 Genetic Analyzer", "AB 3130xL Genetic Analyzer", "AB 3130 Genetic Analyzer", "AB 310 Genetic Analyzer", "unspecified" )
}

FILETYPES = (
"bam", "sra", "kar", "srf", "sff", "fastq", "tab", "454_native",
"Helicos_native", "SOLiD_native_csfasta", "SOLiD_native_qual", "SOLiD_native",
"PacBio_HDF5", "CompleteGenomics_native", "bam_header", "reference_fasta"
)


### Create Instructions Page ###
worksheet = workbook.add_worksheet('Instructions and Contact Info')
worksheet.set_tab_color('red')
worksheet.set_column('A:A', 45)
worksheet.set_column('B:D',25)
worksheet.set_row('35:35', 32)

### Write Instructions on Page ###
worksheet.write_column('A1', ['submission_name', 'contact_name', 'inform_on_status'],contact)
worksheet.write_comment('A1', 'Must be a unique name for the submitting user or group.  Is a tracking tool but not a title for users.')
worksheet.write_comment('A2', 'Name of submission owner')
worksheet.write_comment('A3', 'Email address for information updates')

worksheet.write_column('B1', ['name of submission here', 'your name here', 'your email address here'], bold)

worksheet.write_column('A5', [
            'Instructions:', 
            'Please make sure you have completed your Study and Sample Registration with dbGaP first.', 
            'All columns with white-on-blue headers are REQUIRED.',
            'Do not delete columns from the sheet.  If unsure of content, either leave blank or contact SRA.', 
            'Columns with white-on-grey headers are OPTIONAL.',
            'Each column that has a red triangle in the upper-right corner has a comment that can be displayed if you hover over the header.',
            'Some column headers have hyperlinks to NCBI webpages.',
            'The YELLOW columns have drop-down menus that allow you to select from a controlled vocabulary. Once specified for one row, these values can be copied-and-pasted down.'
            ], bold)
worksheet.insert_image('A14', '/home/stineaj/bin/resources/example.png', {'x_scale': 0.9, 'y_scale': 0.9, 'x_offset': 15})
worksheet.write_column('A31', [
            'Many of the columns also have data checks - if you received a warning, please verify that you have attempted to enter a correct value.', 
            'NOTE: There are data checks and autocomplete features in this spreadsheet that are not compatible with Libre- and Open-Office. If you use one of these suites, please manually consult the platform and instrument information on the last page.'
            ], bold)
            
worksheet.write('A34', 'Header key:', bold)
worksheet.write('A35', 'red triangles indicate pop-up comments for that field', paired)
worksheet.write_comment('A35', 'Like this one')
worksheet.write('B35', 'required for ALL data types', required)
worksheet.write('C35', 'required for aligned data', aligned)
worksheet.write('D35', 'paired-end data only', paired)
worksheet.write('A37', 'SRA submission overview:', bold)
worksheet.write('B37', 'http://www.ncbi.nlm.nih.gov/books/NBK242619/', url_format)


### Create Terms Page ### 
worksheet = workbook.add_worksheet('Terms')
worksheet.set_column('A:J', 25)
worksheet.add_table('A2:B30', {'style': 'Table Style Light 9', 'autofilter': False, 'data': STRATEGY, 'columns': [{'header': 'Strategy'},{'header': 'Description'},]})
worksheet.add_table('A32:B39', {'style': 'Table Style Light 9', 'autofilter': False, 'data': SOURCE, 'columns': [{'header': 'Source'},{'header': 'Description'},]})
worksheet.add_table('A41:B68', {'style': 'Table Style Light 9', 'autofilter': False, 'data': SELECTION, 'columns': [{'header': 'Selection'},{'header': 'Description'},]})
worksheet.write_row('A71',('platforms', 'ILLUMINA', '_LS454', 'COMPLETE_GENOMICS', 'ABI_SOLID', 'PACBIO_SMRT', 'ION_TORRENT', 'CAPILLARY', 'OXFORD_NANOPORE', 'HELICOS'), bold)
worksheet.write_column('A72',('ILLUMINA', '_LS454', 'COMPLETE_GENOMICS', 'ABI_SOLID', 'PACBIO_SMRT', 'ION_TORRENT', 'CAPILLARY', 'OXFORD_NANOPORE', 'HELICOS'))
worksheet.write_column('B72', PLATFORMS['ILLUMINA'])
worksheet.write_column('C72', PLATFORMS['_LS454'])
worksheet.write('D72', 'COMPLETE_GENOMICS')
worksheet.write_column('E72', PLATFORMS['ABI_SOLID'])
worksheet.write_column('F72', PLATFORMS['PACBIO_SMRT'])
worksheet.write_column('G72', PLATFORMS['ION_TORRENT'])
worksheet.write_column('H72', PLATFORMS['CAPILLARY'])
worksheet.write_column('I72', PLATFORMS['OXFORD_NANOPORE'])
worksheet.write_column('J72', PLATFORMS['HELICOS'])


### Define Names for Platforms and Models ###
workbook.define_name('Strategy',        '=Terms!$A$3:$A$30')
workbook.define_name('Source',          '=Terms!$A$33:$A$39')
workbook.define_name('Selection',       '=Terms!$A$42:$A$68')
workbook.define_name('Platforms',       '=Terms!$A$72:$A$80')
workbook.define_name('ILLUMINA',        '=Terms!$B$72:$B$84')
workbook.define_name('_LS454',          '=Terms!$C$72:$C$78')
workbook.define_name('COMPLETE_GENOMICS', '=Terms!$D$72:$D$72')
workbook.define_name('ABI_SOLID',       '=Terms!$E$72:$E$82')
workbook.define_name('PACBIO_SMRT',     '=Terms!$F$72:$F$74')
workbook.define_name('ION_TORRENT',     '=Terms!$G$72:$G$74')
workbook.define_name('CAPILLARY',       '=Terms!$H$72:$H$79')
workbook.define_name('OXFORD_NANOPORE', '=Terms!$I$72:$I$74')
workbook.define_name('HELICOS',         '=Terms!$J$72:$J$73')


### Create Data Page ###
worksheet = workbook.add_worksheet('SRA_Data')
worksheet.set_tab_color('yellow')


### Write links from Library Controlled Vocab to the Terms page###
worksheet.write_url('E1', 'internal:Terms!$A$3')
worksheet.write_url('F1', 'internal:Terms!$A$33')
worksheet.write_url('G1', 'internal:Terms!$A$42')
worksheet.write_url('I1', 'internal:Terms!$A$71')


### Format Data Page and Add Headers ###
worksheet.set_column('A:U', 20)
worksheet.freeze_panes(0,3)
worksheet.write_row('A1', ('phs_accession', 'sample_name', 'library_ID', 'title/short description', 'library_strategy (click for details)', 'library_source (click for details)', 'library_selection (click for details)', 'library_layout', 'platform (click for details)', 'instrument_model', 'design_description'), required)
worksheet.write_row('L1', ('reference_genome_assembly (or accession)', 'alignment_software'), aligned)
worksheet.write('N1', 'forward_read_length', required)
worksheet.write('O1', 'reverse_read_length', paired)
worksheet.write_row('P1', ('filetype', 'filename', 'MD5_checksum'), required)
worksheet.write_row('S1', ('filetype', 'filename', 'MD5_checksum'), paired)

### Add Data Page Comments ###
worksheet.write_comment('A1', 'The phs accession in the format phs000000; NOT including version (.v/.p) numbers.')
worksheet.write_comment('B1', 'The sample_name as described in the dbGaP submission documents (also called submitted_sample_name or SAMPID).')
worksheet.write_comment('C1', 'A short unique identifier for the sequencing library. Each library_ID MUST be unique!')
worksheet.write_comment('D1', 'Short description that will identify the dataset on public pages. A clear and concise formula for the title would be:' + "\n\n" + '{methodology}' + "\n\n" + 'of {organism}: {sample info}' + "\n\n" + ' e.g.' + "\n\n" + 'RNA-Seq of mus musculus: adult female spleen', {'width': 300, 'height': 200})
worksheet.write_comment('H1', 'Paired-end or Single')
worksheet.write_comment('K1', 'Free-form description of the methods used to create the sequencing library; a brief \'materials and methods\' section.')
worksheet.write_comment('L1', 'For bam-format files. Please include NCBI accession(s) or assembly name.')
worksheet.write_comment('M1', 'Please include version #, if known.')
worksheet.write_comment('N1', 'Maximum length of the first biological read in a paired library or the only read of a fragment library.')
worksheet.write_comment('O1', 'PAIRED ONLY')
worksheet.write_comment('P1', 'Format of the data file to be submitted.  Must be one of the types in the list. ')
worksheet.write_comment('Q1', 'File name including all extensions, but NOT path information.')
worksheet.write_comment('R1', 'Checksum generated by the MD5 algorithm for the indicated file.')
worksheet.write_comment('S1', 'PAIRED ONLY Format of the data file to be submitted.  Must be one of the types in the list.')
worksheet.write_comment('T1', 'PAIRED ONLY File name including all extensions, but NOT path information')
worksheet.write_comment('U1', 'PAIRED ONLY Checksum generated by the MD5 algorithm for the indicated file')

###  Enter Samples from the dbGaP service into rows along with formatting. ###
x = 2
for node in tree.iter('Sample'):
   ssid = node.attrib.get('submitted_sample_id')
   status = node.attrib.get('dbgap_status')
   if status == 'Loaded':
       worksheet.write_row('E' + str(x),(None,None,None,None,None,None), lists)
       worksheet.write('A' + str(x), phs, bold)
       worksheet.write('B' + str(x), ssid, bold)
       worksheet.data_validation('E' + str(x), {'validate': 'list', 'source': '=Strategy'})
       worksheet.data_validation('F' + str(x), {'validate': 'list', 'source': '=Source'})
       worksheet.data_validation('G' + str(x), {'validate': 'list', 'source': '=Selection'})
       worksheet.data_validation('H' + str(x), {'validate': 'list', 'source': ['single', 'paired']})
       worksheet.data_validation('I' + str(x), {'validate': 'list', 'source': '=Platforms'})
       worksheet.data_validation('J' + str(x), {'validate': 'list', 'source': '=INDIRECT($I' + str(x) + ')'})
       x += 1
   print '  %s :: %s :: %s' % (ssid, status, 'A' + str(x))


### Close Out Workbook ###   
workbook.close()
