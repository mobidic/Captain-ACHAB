#!/usr/bin/perl

##### vcfToExcel.pl ####

# Auteur : Thomas Guignard 2016
# USAGE : vcfToExcel.pl --vcf <vcf_file> --cas <index_sample_name> --pere <father_sample_name> --mere <mother_sample_name> --control <control_sample_name>  --output <txt_file> --caller <freebayes|GATK> --trio <YES|NO> --candidats <file with gene symbol of interest>  --geneSummary <file with format geneSymbol\tsummary>  --pLIFile <file containing gene symbol and pLI from ExAC database>
#
#ExAC pLI datas ftp://ftp.broadinstitute.org/pub/ExAC_release/release1/functional_gene_constraint/

# Description : 
# Create an User friendly Excel file from an annotated VCF file. 

# Version : 
# v1.0.0 20161020 Initial Implementation 
# v6 20171201 Use VCF header to get parameter

use strict; 
use warnings;
use Getopt::Long; 
use Pod::Usage;
use List::Util qw(first);
use Excel::Writer::XLSX;

my $man = 0;
my $help = 0;
my $current_line;
my $incfile;
my $outfile;
my $cas = "";
my $mere = "";
my $control = "";
my $pere = "";
my $caller = "";
my $trio = "";
my $candidates = "";
my @candidatesList;
my $candidates_line;

my $geneSummary="";
my $geneSummary_Line;
my @geneSummaryList;
my $geneSummaryConcat;

my $pLIFile = "";
my $pLI_Line;
my @pLI_List;
my $pLI_values;  

my @line;
my @unorderedLine;
my @orderedLine;

my $nbrSample = 0;
my $indexPere = 0;
my $indexMere = 0;
my $indexCas = 0;
my $indexControl = 0;


#$arguments = GetOptions( "vcf=s" => \$incfile, "output=s" => \$outfile, "cas=s" => \$cas, "pere=s" => \$pere, "mere=s" => \$mere, "control=s" => \$control );
#$arguments = GetOptions( "vcf=s" => \$incfile ) or pod2usage(-vcf => "$0: argument required\n") ;

GetOptions( 	"vcf=s" => \$incfile,
			 	"output=s" => \$outfile,
				"cas=s" => \$cas,
				"pere=s" => \$pere, 
				"mere=s" => \$mere, 
				"control=s" => \$control,
				"caller=s" => \$caller,
				"trio=s" => \$trio,
				"candidats=s" => \$candidates,
				"geneSummary=s" => \$geneSummary,
        "pLIFile=s" => \$pLIFile,
				"man" => \$man,
				"help" => \$help);


#check mandatory arguments



			 
print  STDERR "Processing vcf file ... \n" ; 


open( VCF , "<$incfile" )or die("Cannot open vcf file $incfile") ;
open(OUT,"| /bin/gzip -c >$outfile".".gz") ;
#open(OUTSHORTSORT,"| /bin/gzip -c >$outfile".".shortNsort.gz") ;
#open(OUTPASStrue1pct,">$outfile".".PASStrue1pct.txt") ;



# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new( $cas."_".$pere."_".$mere."_".$control.'.xlsx' );

#create color background for pLI values
my $format_pLI = $workbook->add_format();
#$format_pLI -> set_pattern();


# Add all worksheets
my $worksheet = $workbook->add_worksheet('ALL');
$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLine = 0;

my $worksheetOMIM = $workbook->add_worksheet('OMIM');
$worksheetOMIM->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineOMIM = 0;

my $worksheetACMG = $workbook->add_worksheet('DS_ACMG');
$worksheetACMG->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineACMG = 0;


my $worksheetClinVarPatho = $workbook->add_worksheet('ClinVar_Patho');  
$worksheetClinVarPatho->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineClinVarPatho = 0;


my $worksheetRAREnonPASS = $workbook->add_worksheet('Rare_nonPASS');
$worksheetRAREnonPASS->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineRAREnonPASS = 0;


$worksheetRAREnonPASS->autofilter('A1:AN1000'); # Add autofilter
$worksheetOMIM->autofilter('A1:AN1000'); # Add autofilter
$worksheetClinVarPatho->autofilter('A1:AN1000'); # Add autofilter
$worksheet->autofilter('A1:AN5000'); # Add autofilter




my $worksheetHTZcompo;
my $worksheetAR; 
my $worksheetSNPmereVsCNVpere ;
my $worksheetSNPpereVsCNVmere;
my $worksheetDENOVO;
my $worksheetCandidats;
my $worksheetDELHMZ;

my $worksheetLineHTZcompo ;
my $worksheetLineAR ; 
my $worksheetLineSNPmereVsCNVpere ;
my $worksheetLineSNPpereVsCNVmere ;
my $worksheetLineDENOVO ;
my $worksheetLineCandidats;
my $worksheetLineDELHMZ;


#create additionnal sheet in trio analysis
if ($trio eq "YES"){
	$worksheetHTZcompo = $workbook->add_worksheet('HTZ_compo');
	$worksheetLineHTZcompo = 0;
	$worksheetHTZcompo->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheetHTZcompo->autofilter('A1:AN1000'); # Add autofilter
  
	$worksheetAR = $workbook->add_worksheet('AR');
	$worksheetLineAR = 0;
	$worksheetAR->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheetAR->autofilter('A1:AN1000'); # Add autofilter
  
	$worksheetSNPmereVsCNVpere = $workbook->add_worksheet('SNVmereVsCNVpere');
	$worksheetLineSNPmereVsCNVpere = 0;
	$worksheetSNPmereVsCNVpere->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheetSNPmereVsCNVpere->autofilter('A1:AN1000'); # Add autofilter
  
	$worksheetSNPpereVsCNVmere = $workbook->add_worksheet('SNVpereVsCNVmere');
	$worksheetLineSNPpereVsCNVmere = 0;
	$worksheetSNPpereVsCNVmere->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheetSNPpereVsCNVmere->autofilter('A1:AN1000'); # Add autofilter
  
	$worksheetDENOVO = $workbook->add_worksheet('DENOVO');
	$worksheetLineDENOVO = 0;
	$worksheetDENOVO->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheetDENOVO->autofilter('A1:AN1000'); # Add autofilter

	$worksheetDELHMZ = $workbook->add_worksheet('DEL_HMZ');
	$worksheetLineDELHMZ = 0;
	$worksheetDELHMZ->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheetDELHMZ->autofilter('A1:AN1000'); # Add autofilter
		

#$worksheetLineHTZcompo ++;
#$worksheetLineAR ++; 
#$worksheetLineSNPmereVsCNVpere ++;
#$worksheetLineSNPpereVsCNVmere ++;
#$worksheetLineDENOVO ++;

}

#create sheet for candidats
if($candidates ne ""){
	open( CANDIDATS , "<$candidates")or die("Cannot open candidates file $candidates") ;
	print  STDERR "Processing candidates file ... \n" ; 
	while( <CANDIDATS> ){
	  	$candidates_line = $_;
		chomp $candidates_line;
		push @candidatesList, $candidates_line;

				
	}
	close(CANDIDATS);
	$worksheetCandidats = $workbook->add_worksheet('Candidats');
	$worksheetCandidats->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheetCandidats->autofilter('A1:AN1000'); # Add autofilter
}


#A reminder of the Parameters mostly present in VCF depending on SNP caller used
#
#my %dicoParam;

#added by perl scripts
 # $dicoParam{"COSMICID"}= "Cosmic identifier";
 # #$dicoParam{"COSMIC"}= "Cosmic identifier";
 # $dicoParam{"CLNORIGIN"}= "Clinvar allele origin";
 # $dicoParam{"CLNSIG"}= "Variant Clinical Significance from Clinvar database";
 # $dicoParam{"OMIM"}= "OMIM Pathologie";
 #
 # #added by Seattle Seq Annotation
 # $dicoParam{"DN"}= "inDbSNP";
 # $dicoParam{"DT"}= "in1000Genomes";
 # $dicoParam{"DA"}= "allelesDBSNP";
 # $dicoParam{"FG"}= "functionGVS";
 # $dicoParam{"FD"}= "functionDBSNP";
 # $dicoParam{"GM"}= "accession";
 # $dicoParam{"GL"}= "geneList";
 # $dicoParam{"AAC"}= "aminoAcids";
 # $dicoParam{"PP"}= "proteinPosition";
 # $dicoParam{"CDP"}= "cDNAPosition";
 # $dicoParam{"PH"}= "polyPhen";
 # $dicoParam{"CP"}= "scorePhastCons";
 # $dicoParam{"CG"}= "consScoreGERP";
 # $dicoParam{"CADD"}= "scoreCADD";
 # $dicoParam{"AA"}= "chimpAllele";
 # $dicoParam{"CN"}= "CNV";
 # $dicoParam{"HA"}= "AfricanHapMapFreq";
 # $dicoParam{"HE"}= "EuropeanHapMapFreq";
 # $dicoParam{"HC"}= "AsianHapMapFreq";
 # $dicoParam{"DG"}= "hasGenotypes";
 # $dicoParam{"DV"}= "dbSNPValidation";
 # $dicoParam{"RM"}= "repeatMasker";
 # $dicoParam{"RT"}= "tandemRepeat";
 # $dicoParam{"CA"}= "clinicalAssociation";
 # $dicoParam{"DSP"}= "distanceToSplice";
 # $dicoParam{"KP"}= "keggPathway";
 # $dicoParam{"CPG"}= "cpgIslands";
 # $dicoParam{"GESP"}= "genomesESP";
 # $dicoParam{"GEXAC"}= "genomesExAC";
 # $dicoParam{"GS"}= "granthamScore";
 # $dicoParam{"MR"}= "microRNAs";
 # 
 #   if ($caller eq "GATK"){
 #
 # # added by GATK
 # $dicoParam{"AC"}= "Allele count in genotypes, for each ALT allele, in the same order as listed";
 # $dicoParam{"AF"}= "Allele Frequency, for each ALT allele, in the same order as listed";
 # $dicoParam{"AN"}= "Total number of alleles in called genotypes";
	#$dicoParam{"BaseQRankSum"}= "Z-score from Wilcoxon rank sum test of Alt Vs. Ref base qualities";
	#$dicoParam{"ClippingRankSum"}="Z-score From Wilcoxon rank sum test of Alt vs. Ref number of hard clipped bases";
	#$dicoParam{"DB"}= "dbSNP Membership";
	#$dicoParam{"DP"}= "Approximate read depth; some reads may have been filtered";
	#$dicoParam{"DS"}= "Were any of the samples downsampled?";
	#$dicoParam{"Dels"}= "Fraction of Reads Containing Spanning Deletions";
	#$dicoParam{"ExcessHet"}="Phred-scaled p-value for exact test of excess heterozygosity";
	#$dicoParam{"FS"}= "Phred-scaled p-value using Fisher's exact test to detect strand bias";
	#$dicoParam{"HaplotypeScore"}= "Consistency of the site with at most two segregating haplotypes";
	#$dicoParam{"InbreedingCoeff"}= "Inbreeding coefficient as estimated from the genotype likelihoods per-sample when compared against the Hardy-Weinberg expectation";
	#$dicoParam{"MLEAC"}= "Maximum likelihood expectation (MLE) for the allele counts (not necessarily the same as the AC), for each ALT allele, in the same order as listed";
	#$dicoParam{"MLEAF"}= "Maximum likelihood expectation (MLE) for the allele frequency (not necessarily the same as the AF), for each ALT allele, in the same order as listed";
	#$dicoParam{"MQ"}= "RMS Mapping Quality";
	#$dicoParam{"MQ0"}= "Total Mapping Quality Zero Reads";
	#$dicoParam{"MQRankSum"}= "Z-score From Wilcoxon rank sum test of Alt vs. Ref read mapping qualities";
	#$dicoParam{"QD"}= "Variant Confidence/Quality by Depth";
	#$dicoParam{"RPA"}= "Number of times tandem repeat unit is repeated, for each allele (including reference)";
	#$dicoParam{"RU"}= "Tandem repeat unit (bases)";
	#$dicoParam{"ReadPosRankSum"}= "Z-score from Wilcoxon rank sum test of Alt vs. Ref read position bias";
	#$dicoParam{"SOR"}= "Symmetric Odds Ratio of 2x2 contingency table to detect strand bias";
	#$dicoParam{"STR"}= "Variant is a short tandem repeat)";
	#$dicoParam{"AB"}="AB Allele balance at heterozygous sites: a number between 0 and 1 representing the ratio of reads showing the reference allele to all reads, considering only reads from     individuals called as heterozygous";
	#$dicoParam{"ABP"}="ABP Allele balance probability at heterozygous sites: Phred-scaled upper-bounds estimate of the probability of observing the deviation between ABR and ABA given E(ABR/A    BA) ~ 0.5, derived using Hoeffding's inequality";
  #
  ##added by GATK HC (from dijex exome)
  #
	#$dicoParam{"ACMG"}="ACMG diseases";
	#$dicoParam{"BATCH_AC"}="Batch allele count in genotypes";
	#$dicoParam{"BATCH_AF"}="Batch allele count in genotypes";
	#$dicoParam{"BATCH_AN"}="Batch total number of alleles in called genotypes, for each ALT allele, in the same order as listed";
	#$dicoParam{"BATCH_GTC"}="Batch genotypes count";
	#$dicoParam{"BATCH_GTN"}="Batch total number of called genotypes";
	#$dicoParam{"BATCH_VARC"}="Batch variant count in genotypes, for each ALT allele, in the same order as listed";
	#$dicoParam{"CTRL_AC"}="Control allele count in genotypes, for each ALT allele, in the same order as listed";
	#$dicoParam{"CTRL_AF"}="Control allele frequency, for each ALT allele, in the same order as listed";
	#$dicoParam{"CTRL_AN"}="Control total number of alleles in called genotypes";
	#$dicoParam{"CTRL_GTC"}="Control genotypes count";
	#$dicoParam{"CTRL_GTN"}="Control total number of called genotypes";
	#$dicoParam{"CTRL_VARC"}="Control variant count in genotypes, for each ALT allele, in the same order as listed";
	#$dicoParam{"RAW_MQ"}="Raw data for RMS Mapping Quality";
	#$dicoParam{"RECESSIVE"}="Recessive";
	#$dicoParam{"DENOVO"}="De novo candidates";
  #
  #
#F#ORMAT freebayes	GT:DP:DPR:RO:QR:AO:QA:GL

#FORMAT GATK		GT:AD:DP:GQ:PL
	#$dicoParam{"GT"}= "Genotype";
	#$dicoParam{"AD"}= "Allelic depths for the ref and alt alleles in the order listed";
	#$dicoParam{"DP"}= "Approximate read depth (reads with MQ 255 or with bad mates are filtered)";
	#$dicoParam{"GQ"}= "Genotype Quality";
	#$dicoParam{"PL"}= "Normalized, Phred-scaled likelihoods for genotypes as defined in the VCF specification";
  #
#F#ILTER BY GATK
	#$dicoParam{"FS200"}= "FS > 200.0";
	#$dicoParam{"LowQual"}= "Low quality";
	#$dicoParam{"QD2"}= "QD < 2.0";
	#$dicoParam{"RPRS-8"}= "ReadPosRankSum < -8.0";
	#$dicoParam{"SnpCluster"}= "SNPs found in clusters";
  #
#} #
  #



#added by freebayes

# if ($caller eq "freebayes"){
				
	#$dicoParam{"NS"}="Number of samples with data";
	#$dicoParam{"DP"}="Total read depth at the locus";
	#$dicoParam{"DPB"}="Total read depth per bp at the locus; bases in reads overlapping / bases in haplotype";
	#$dicoParam{"AC"}="Total number of alternate alleles in called genotypes";
	#$dicoParam{"AN"}="Total number of alleles in called genotypes";
	#$dicoParam{"AF"}="Estimated allele frequency in the range (0,1]";
	#		
	#$dicoParam{"RO"}="RO Reference allele observation count, with partial observations recorded fractionally";
	#$dicoParam{"AO"}="AO Alternate allele observations, with partial observations recorded fractionally";
	#$dicoParam{"PRO"}="PRO Reference allele observation count, with partial observations recorded fractionally";
	#$dicoParam{"PAO"}="PAO Alternate allele observations, with partial observations recorded fractionally";
	#$dicoParam{"QR"}="QR Reference allele quality sum in phred";
	#$dicoParam{"QA"}="QA Alternate allele quality sum in phred";
	#$dicoParam{"PQR"}="PQR Reference allele quality sum in phred for partial observations";
	#$dicoParam{"PQA"}="PQA Alternate allele quality sum in phred for partial observations";
	#$dicoParam{"SRF"}="SRF Number of reference observations on the forward strand";
	#$dicoParam{"SRR"}="SRR Number of reference observations on the reverse strand";
	#$dicoParam{"SRP"}="SRP Strand balance probability for the reference allele: Phred-scaled upper-bounds estimate of the probability of observing the deviation between SRF and SRR given E(SRF/SRR) ~ 0.5, derived using Hoeffding's inequality";
	#$dicoParam{"AB"}="AB Allele balance at heterozygous sites: a number between 0 and 1 representing the ratio of reads showing the reference allele to all reads, considering only reads from individuals called as heterozygous";
	#$dicoParam{"ABP"}="ABP Allele balance probability at heterozygous sites: Phred-scaled upper-bounds estimate of the probability of observing the deviation between ABR and ABA given E(ABR/ABA) ~ 0.5, derived using Hoeffding's inequality";
	#$dicoParam{"RPL"}="RPL Reads Placed Left: number of reads supporting the alternate balanced to the left (5') of the alternate allele";
	#$dicoParam{"RPR"}="RPR Reads Placed Right: number of reads supporting the alternate balanced to the right (3') of the alternate allele";
	#$dicoParam{"SRP"}="SRP Strand balance probability for the reference allele: Phred-scaled upper-bounds estimate of the probability of observing the deviation between SRF and SRR given E(SRF/SRR) ~ 0.5, derived using Hoeffding's inequality";
	#$dicoParam{"SAP"}="SAP Strand balance probability for the alternate allele: Phred-scaled upper-bounds estimate of the probability of observing the deviation between SAF and SAR given E(SAF/SAR) ~ 0.5, derived using Hoeffding's inequality";
	#$dicoParam{"SAF"}="SAF Number of alternate observations on the forward strand";
	#$dicoParam{"SAR"}="SAR Number of alternate observations on the reverse strand";                                                            
	#$dicoParam{"RUN"}="Run length: the number of consecutive repeats of the alternate allele in the reference genome";
	#$dicoParam{"RPP"}="RPP Read Placement Probability: Phred-scaled upper-bounds estimate of the probability of observing the deviation between RPL and RPR given E(RPL/RPR) ~ 0.5, derived using Hoeffding's inequality";
	#$dicoParam{"RPPR"}="RPPR Read Placement Probability for reference observations: Phred-scaled upper-bounds estimate of the probability of observing the deviation between RPL and RPR given E(RPL/RPR) ~ 0.5, derived using Hoeffding's inequality";                                    
	#$dicoParam{"EPP"}="EPP End Placement Probability: Phred-scaled upper-bounds estimate of the probability of observing the deviation between EL and ER given E(EL/ER) ~ 0.5, derived using Hoeffding's inequality";                                  
	#$dicoParam{"EPPR"}="EPPR End Placement Probability for reference observations: Phred-scaled upper-bounds estimate of the probability of observing the deviation between EL and ER given E(EL/ER) ~ 0.5, derived using Hoeffding's inequality";
	#$dicoParam{"DPRA"}="Alternate allele depth ratio.  Ratio between depth in samples with each called alternate allele and those without.";
	#$dicoParam{"ODDS"}="Log odds ratio of the best genotype combination to the second-best.";
	#$dicoParam{"GTI"}="Number of genotyping iterations required to reach convergence or bailout.";
	#$dicoParam{"TYPE"}="Type of allele, either snp, mnp, ins, del, or complex.";
	#$dicoParam{"CIGAR"}="The extended CIGAR representation of each alternate allele, with the exception that '=' is replaced by 'M' to ease VCF parsing.  Note that INDEL alleles do not have the first matched base (which is provided by default, per the spec) referred to by the CIGAR.";
	#$dicoParam{"NUMALT"}="Number of unique non-reference alleles in called genotypes at this position.";
	#$dicoParam{"MEANALT"}="Mean number of unique non-reference allele observations per sample with the corresponding alternate alleles.";
	#$dicoParam{"LEN"}="allele length";           
	#$dicoParam{"MQM"}="Mean mapping quality of observed alternate alleles, from 0 to 255";
	#$dicoParam{"MQMR"}="Mean mapping quality of observed reference alleles,from 0 to 255";
	#$dicoParam{"PAIRED"}="Proportion of observed alternate alleles which are supported by properly paired read fragments";
	#$dicoParam{"PAIREDR"}="Proportion of observed reference alleles which are supported by properly paired read fragments";
  #                                                                              
  #
#FORMAT GATK		GT:AD:DP:GQ:PL
#FORMAT freebayes	GT:DP:DPR:RO:QR:AO:QA:GL

  #$dicoParam{"GT"}="Genotype";
#               $dicoParam{"GQ"}="Genotype Quality, the Phred-scaled marginal (or unconditional) probability of the called genotype";
#               $dicoParam{"GL"}="Genotype Likelihood, log10-scaled likelihoods of the data given the called genotype for each possible genotype generated from the reference and alternate alleles given the sample ploidy";
#               $dicoParam{"DP"}="Read Depth";
#               $dicoParam{"DPR"}="Number of observation for each allele";
#               $dicoParam{"RO"}="Reference allele observation count";
#               $dicoParam{"QR"}="Sum of quality of the reference observations";
#               $dicoParam{"AO"}="Alternate allele observation count";
#               $dicoParam{"QA"}="Sum of quality of the alternate observations";

#}

#create dico of functions for SNV and indel
my %dicoFonction;
$dicoFonction{"coding-unknown"} = 0 ;
$dicoFonction{"coding-unknown-near-splice"} = 0 ;
$dicoFonction{"missense"} = 0 ;
$dicoFonction{"missense-near-splice"} = 0 ;
$dicoFonction{"splice-acceptor"} = 0 ;
$dicoFonction{"splice-donor"} = 0 ;
$dicoFonction{"stop-gained"} = 0 ;
$dicoFonction{"stop-gained-near-splice"} = 0 ;
$dicoFonction{"stop-lost"} = 0 ;
$dicoFonction{"stop-lost-near-splice"} = 0 ;
$dicoFonction{"coding"} = 0 ;
$dicoFonction{"coding-near-splice"} = 0 ;
$dicoFonction{"codingComplex"} = 0 ;
$dicoFonction{"codingComplex-near-splice"} = 0 ;
$dicoFonction{"frameshift"} = 0 ;
$dicoFonction{"synonymous-near-splice"} = 0 ;
$dicoFonction{"intron-near-splice"} = 0 ;
$dicoFonction{"frameshift-near-splice"} = 0 ;




#List of ACMG incidentalome genes
my @ACMGlist = ("ACTA2","ACTC1","APC","APOB","ATP7B","BMPR1A","BRCA1","BRCA2","CACNA1S","COL3A1","DSC2","DSG2","DSP","FBN1","GLA","KCNH2","KCNQ1","LDLR","LMNA","MEN1","MLH1","MSH2","MSH6","MUTYH","MYBPC3","MYH11","MYH7","MYL2","MYL3","NF2","OTC","PCSK9","PKP2","PMS2","PRKAG2","PTEN","RB1","RET","RYR1","RYR2","SCN5A","SDHAF2","SDHB","SDHC","SDHD","SMAD3","SMAD4","STK11","TGFBR1","TGFBR2","TMEM43","TNNI3","TNNT2","TP53","TPM1","TSC1","TSC2","VHL","WT1");
                  
#empty line to erase false HTZ composite lines
my @emptyArray = (" "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "); 


#create dico of gene summary from refseq
my %dicoGeneSummary;

if($geneSummary ne ""){
	open( GENESUMMARY , "<$geneSummary")or die("Cannot open geneSummary file $geneSummary") ;
	
	print  STDERR "Processing geneSummary file ... \n" ; 
	
	while( <GENESUMMARY> ){
	  	$geneSummary_Line = $_;
		chomp $geneSummary_Line;
		@geneSummaryList = split( /\t/, $geneSummary_Line);	
		$dicoGeneSummary{$geneSummaryList[0]} = $geneSummaryList[1];
				
	}
	close(GENESUMMARY);
}




#create dico of ExAC pLI data / gene
my %dicopLI;

my $pLI_Comment = "pLI - the probability of being loss-of-function intolerant (intolerant of both heterozygous and homozygous lof variants)\npRec - the probability of being intolerant of homozygous, but not heterozygous lof variants\npNull - the probability of being tolerant of both heterozygous and homozygous lof variants";


if($pLIFile ne ""){
	open( PLIFILE , "<$pLIFile")or die("Cannot open pLI file $pLIFile") ;
	
	print  STDERR "Processing pLI file ... \n" ; 
	
	while( <PLIFILE> ){
	  	$pLI_Line = $_;
		next if ($pLI_Line=~/^transcript/);
		chomp $pLI_Line;

		@pLI_List = split( /\t/, $pLI_Line);	
		$dicopLI{$pLI_List[1]}{"comment"} = "pLI = ".$pLI_List[19]."\npRec = ".$pLI_List[20]."\npNull = ".$pLI_List[21]."\nmis_Zscore = ".$pLI_List[18];
		$dicopLI{$pLI_List[1]}{"pLI"} = $pLI_List[19];
		$dicopLI{$pLI_List[1]}{"pRec"} = $pLI_List[20];
		$dicopLI{$pLI_List[1]}{"pNull"} = $pLI_List[21];

		$dicopLI{$pLI_List[1]}{"color_format"} = sprintf('#%2.2X%2.2X%2.2X',($dicopLI{$pLI_List[1]}{"pLI"}*255 + $dicopLI{$pLI_List[1]}{"pRec"}*255),($dicopLI{$pLI_List[1]}{"pRec"}*255 + $dicopLI{$pLI_List[1]}{"pNull"} * 255),0);
		#print "fromat:".$dicopLI{$pLI_List[1]}{"color_format"}."\t".($dicopLI{$pLI_List[1]}{"pLI"}*255 + $dicopLI{$pLI_List[1]}{"pRec"}*255)."\t".($dicopLI{$pLI_List[1]}{"pRec"}*255 + $dicopLI{$pLI_List[1]}{"pNull"} * 255)."\t".$dicopLI{$pLI_List[1]}{"pLI"}."\t".$dicopLI{$pLI_List[1]}{"pRec"}."\t".$dicopLI{$pLI_List[1]}{"pNull"}."\n";
				
	}
	close(PLIFILE);
}






#Parse VCF header to fill the dictionnary of parameters
my %dicoParam;

while( <VCF> ){
  	$current_line = $_;
		
      
    #filling dicoParam with VCF header INFO and FORMAT 

    if ($current_line=~/^##/){


		  unless ($current_line=~/Description=/){ next }
			chomp $current_line;
      #DEBUG print STDERR "Header line\n";

      if ($current_line =~ /ID=(.+?),.*?Description="(.+?)"/){
    
          $dicoParam{$1}= $2;
          print STDERR "info : ". $1 . "\tdescription: ". $2."\n";
			

			    next;
      
      }else {print STDERR "pattern non reconnu dans cette ligne: ".$current_line ."\n";next} 
		
    }else {last}

}
close(VCF);




#DEBUG
#foreach my $ptp (keys %dicoParam){
#  print STDERR $ptp."\t".$dicoParam{$ptp}."\n";
#}


#counter for shifting of columns according to nb individu
my $cmpt = 0;

#create Dico of sorted columns for a userfriendly output
my %dicoOrderColumns;
$dicoOrderColumns{1}{'colName'} = "#CHROM" ;
$dicoOrderColumns{2}{'colName'} = "POS";
$dicoOrderColumns{3}{'colName'} = "ID" ;
$dicoOrderColumns{4}{'colName'} = "REF";
$dicoOrderColumns{5}{'colName'} = "ALT";
$dicoOrderColumns{6}{'colName'} = "QUAL";
$dicoOrderColumns{7}{'colName'} = "FILTER" ;
$dicoOrderColumns{8}{'colName'} = $cas ;
if($pere ne "" && $mere ne "" && $control ne ""){
	$cmpt = 6;
	$dicoOrderColumns{9}{'colName'} = $pere ;
	$dicoOrderColumns{10}{'colName'} = $mere ;
	$dicoOrderColumns{11}{'colName'} = $control ;
	
	$dicoOrderColumns{12}{'colName'} = "Genotype-".$cas ;
	$dicoOrderColumns{13}{'colName'} = "Genotype-".$pere;
	$dicoOrderColumns{14}{'colName'} = "Genotype-".$mere ;
	$dicoOrderColumns{15}{'colName'} = "Genotype-".$control ;
	


}else{
	if( $mere ne "" && $pere ne ""){
		$cmpt = 4;
		$dicoOrderColumns{9}{'colName'} = $pere ;
		$dicoOrderColumns{10}{'colName'} = $mere ;
		
		$dicoOrderColumns{11}{'colName'} = "Genotype-".$cas ;
		$dicoOrderColumns{12}{'colName'} = "Genotype-".$pere ;
		$dicoOrderColumns{13}{'colName'} = "Genotype-".$mere ;

	}else{
		if( $pere ne ""){
			$cmpt = 2;
			$dicoOrderColumns{9}{'colName'} = $pere ;

			$dicoOrderColumns{10}{'colName'} = "Genotype-".$cas ;
			$dicoOrderColumns{11}{'colName'} = "Genotype-".$pere ;
		}
		if( $mere ne ""){
			$cmpt= 2;
			$dicoOrderColumns{9}{'colName'} = $mere ;
			
			$dicoOrderColumns{10}{'colName'} = "Genotype-".$cas ;
			$dicoOrderColumns{11}{'colName'} = "Genotype-".$mere ;
		}
	
	}
}


#$dicoOrderColumns{9+$cmpt}{'colName'} = "Genotype-".$cas ;
$dicoOrderColumns{10+$cmpt}{'colName'} = "geneList" ;
$dicoOrderColumns{11+$cmpt}{'colName'} = $dicoParam{"OMIM"} ;
$dicoOrderColumns{12+$cmpt}{'colName'} = $dicoParam{"FD"} ;
$dicoOrderColumns{13+$cmpt}{'colName'} = $dicoParam{"FG"} ;
$dicoOrderColumns{14+$cmpt}{'colName'} = $dicoParam{"CDP"} ;
$dicoOrderColumns{15+$cmpt}{'colName'} = $dicoParam{"PP"} ;
$dicoOrderColumns{16+$cmpt}{'colName'} = $dicoParam{"AAC"} ;
$dicoOrderColumns{17+$cmpt}{'colName'} = $dicoParam{"DSP"} ;
$dicoOrderColumns{18+$cmpt}{'colName'} = $dicoParam{"GM"} ;
$dicoOrderColumns{19+$cmpt}{'colName'} = "freqExAC(%)" ;
$dicoOrderColumns{20+$cmpt}{'colName'} = $dicoParam{"GEXAC"} ;
$dicoOrderColumns{21+$cmpt}{'colName'} = "Lien Exac" ;
$dicoOrderColumns{22+$cmpt}{'colName'} = "freqGESP(%)" ;
$dicoOrderColumns{23+$cmpt}{'colName'} = $dicoParam{"GESP"} ;
$dicoOrderColumns{24+$cmpt}{'colName'} = $dicoParam{"CADD"} ;
$dicoOrderColumns{25+$cmpt}{'colName'} = $dicoParam{"PH"} ;
$dicoOrderColumns{26+$cmpt}{'colName'} = $dicoParam{"CP"} ;
$dicoOrderColumns{27+$cmpt}{'colName'} = $dicoParam{"CG"} ;
$dicoOrderColumns{28+$cmpt}{'colName'} = $dicoParam{"KP"} ;
$dicoOrderColumns{29+$cmpt}{'colName'} = $dicoParam{"DA"} ;
$dicoOrderColumns{30+$cmpt}{'colName'} = $dicoParam{"DN"} ;

#BUG with CA clinicalAssociation from seattle seq, which is absent from vcf in my test 
#$dicoOrderColumns{31+$cmpt}{'colName'} = $dicoParam{"CA"} ;
#$dicoOrderColumns{32+$cmpt}{'colName'} = $dicoParam{"CLNORIGIN"} ;
#$dicoOrderColumns{33+$cmpt}{'colName'} = $dicoParam{"CLNSIG"} ;
##$dicoOrderColumns{34+$cmpt}{'colName'} = $dicoParam{"COSMIC"} ;
#$dicoOrderColumns{34+$cmpt}{'colName'} = $dicoParam{"COSMICID"} ;
#$dicoOrderColumns{35+$cmpt}{'colName'} = "NS/SS/I" ;

$dicoOrderColumns{31+$cmpt}{'colName'} = $dicoParam{"CLNORIGIN"} ;
$dicoOrderColumns{32+$cmpt}{'colName'} = $dicoParam{"CLNSIG"} ;
#$dicoOrderColumns{34+$cmpt}{'colName'} = $dicoParam{"COSMIC"} ;
$dicoOrderColumns{33+$cmpt}{'colName'} = $dicoParam{"COSMICID"} ;
$dicoOrderColumns{34+$cmpt}{'colName'} = "NS/SS/I" ;



#DEBUG this print is useful to check missing parameter in the VCF
#foreach (sort {$a <=> $b} keys %dicoOrderColumns){
# 	print $_."\t".$dicoOrderColumns{$_}{'colName'}."\n";
#}                  



my %dicoGeneForHTZcompo;
my $previousGene ="";

my $indexUnSort = 0;

#Start parsing VCF

open( VCF , "<$incfile" )or die("Cannot open vcf file $incfile") ;

my @colonnes = ("#CHROM","POS","ID","REF","ALT","QUAL","FILTER");

while( <VCF> ){
  	$current_line = $_;

    # print STDERR $current_line."\n";

  next if ($current_line=~/^##/);

	chomp $current_line;
	
	my $freqGESP= -1;
	my $freqEXAC= -1;
	my $nsssi = "";

	my %genoCas;
	my %genoCont;
	my %genoMere;
	my %genoPere;


	if ( $current_line =~ m/^#CHROM/g )   {
		@line = split( /\t/, $current_line );	
		for( my $i = 9; $i < scalar @line; $i++){
			push @colonnes, $line[$i];
			$nbrSample ++ ;
		#	print "",$i-2,"\t",$line[$i],"\n";

		}
		for( my $j = 9; $j < scalar @line; $j++){
			push @colonnes, "Genotype-".$line[$j];
	#		print "Genotype-".$line[$j],"\n";
			if (defined $cas && $cas eq $line[$j]){
				$indexCas= $j;
			}
      if (defined $pere && $pere eq $line[$j]){	
				$indexPere= $j;
      }
      if (defined $mere && $mere eq $line[$j]){
				$indexMere= $j;
      }
      if (defined $control && $control eq $line[$j]){
				$indexControl= $j;
      }

			
		}
    


		foreach my $param (sort keys %dicoParam) {
			push @colonnes, $dicoParam{$param};
      #DEBUG
      #print STDERR "toto      ".$dicoParam{$param}."\t".$param."\n";
		}

#complete ordered list of param 
#    for (my $i = 0 ; $i <= scalar @listParam; $i ++){
#      foreach my $param ( keys %{$listParam{$i}} ) {
#        push @colonnes, $listParam{$i}{$param};  
#      }
#    }


		my $chaine="";		
		my $orderedChaine="";

		foreach my $col (@colonnes) {
			$chaine .= $col."\t"; 
		}
		
#		$chaine .= "freqGESP(%)\tNS/SS/I\trecessive\tde novo\tLien Exac\tfreqExAC\n";		
		$chaine .= "freqGESP(%)\tNS/SS/I\tLien Exac\tfreqExAC(%)\n";		


		
		print OUT $chaine;
    #DEBUG
    #print STDERR $chaine."\n";

		chomp $chaine;

		#Sort chaine to get correct index
    		@unorderedLine = split( /\t/, $chaine );

    #pour chaque nom de colonne...
		foreach my $colindex (@unorderedLine){
      #DEBUG
      #print $colindex."   colindex\n";
			
      #...compare chaque nom présent dans le dico ordonné poru obtenir le bon index pour la sortie en ordre
      foreach my $finCol (keys %dicoOrderColumns){
        #DEBUG
        #print $colindex ."\t".$dicoOrderColumns{$finCol}{'colName'}."\n";
				
        if ($dicoOrderColumns{$finCol}{'colName'} eq $colindex ){
					$dicoOrderColumns{$finCol}{'colIndex'} =  $indexUnSort;				
					last;
				}
			}
			$indexUnSort ++;
		}

    #Create ordered string
		foreach my $finalcol ( sort {$a <=> $b}  (keys %dicoOrderColumns) ) {

			#print "ordered\t".$finalcol."\t".$unorderedLine[$dicoOrderColumns{$finalcol}{'colIndex'}]."\n"; 
			$orderedChaine .= $unorderedLine[$dicoOrderColumns{$finalcol}{'colIndex'}]."\t";

		}
		
		$orderedChaine .= "\n";

#		print OUTSHORTSORT $orderedChaine;
#		print OUTPASStrue1pct $orderedChaine;

		    @orderedLine = split( /\t/, $orderedChaine );
	
	    	$worksheet->write_row( 0, 0, \@orderedLine );
	    	$worksheetOMIM->write_row( 0, 0, \@orderedLine );
	    	$worksheetACMG->write_row( 0, 0, \@orderedLine );
	    	$worksheetClinVarPatho->write_row( 0, 0, \@orderedLine );
	    	$worksheetRAREnonPASS->write_row( 0, 0, \@orderedLine );
		
        #write comment for pLI
	    	$worksheet->write_comment( 0, 9+$cmpt, $pLI_Comment,  x_scale => 3 );
	    	$worksheetOMIM->write_comment( 0, 9+$cmpt, $pLI_Comment,  x_scale => 3  );
	    	$worksheetACMG->write_comment( 0, 9+$cmpt, $pLI_Comment,  x_scale => 3  );


	    	$worksheetLine ++;
	    	$worksheetLineOMIM ++;
	    	$worksheetLineACMG ++;



		$worksheetClinVarPatho->write_comment( 0, 9+$cmpt, $pLI_Comment,  x_scale => 3  );
		$worksheetRAREnonPASS->write_comment( 0, 9+$cmpt, $pLI_Comment,  x_scale => 3  );

	    	$worksheetLineClinVarPatho ++;
	    	$worksheetLineRAREnonPASS ++;



		if ($trio eq "YES"){
			$worksheetHTZcompo->write_row( 0, 0, \@orderedLine );
			$worksheetAR->write_row( 0, 0, \@orderedLine );
			$worksheetSNPmereVsCNVpere->write_row( 0, 0, \@orderedLine );
			$worksheetSNPpereVsCNVmere->write_row( 0, 0, \@orderedLine );
			$worksheetDENOVO->write_row( 0, 0, \@orderedLine );

			$worksheetDELHMZ->write_row( 0, 0, \@orderedLine );
			
			#write pLI comment
			$worksheetHTZcompo->write_comment(0, 9+$cmpt, $pLI_Comment,  x_scale => 3);
			$worksheetAR->write_comment( 0, 9+$cmpt, $pLI_Comment,  x_scale => 3);
			$worksheetSNPmereVsCNVpere->write_comment( 0, 9+$cmpt, $pLI_Comment,  x_scale => 3);
			$worksheetSNPpereVsCNVmere->write_comment(0, 9+$cmpt, $pLI_Comment,  x_scale => 3);
			$worksheetDENOVO->write_comment( 0, 9+$cmpt, $pLI_Comment,  x_scale => 3);

			$worksheetDELHMZ->write_comment( 0, 9+$cmpt, $pLI_Comment,  x_scale => 3);


			$worksheetLineHTZcompo ++;
			$worksheetLineHTZcompo ++;
			$worksheetLineAR ++; 
			$worksheetLineSNPmereVsCNVpere ++;
			$worksheetLineSNPpereVsCNVmere ++;
			$worksheetLineDENOVO ++;

			$worksheetLineDELHMZ ++;
		}

	    	if($candidates ne ""){
			$worksheetCandidats->write_row( 0, 0, \@orderedLine );
			$worksheetCandidats->write_comment(0, 9+$cmpt, $pLI_Comment,  x_scale => 3 );
			$worksheetLineCandidats ++;
		}
#		print $chaine;

		
	}else {
		my $chaine="";
		my $orderedChaine="";

		my $alt="";
		my $ref="";
		my %data;

    @line = split( /\t/, $current_line );   
					
		#DEBUG		print $current_line,"\n";	
				
		for( my $i = 0; $i < scalar @line; $i++){
			if ($i < 7){
				$data{$i} = $line[$i];
				if ($i == 3){
					$ref = $line[$i];
				} 	
				if ($i == 4){
					$alt = $line[$i];
				} 	
			}
			#calculate population frequency
			if ($i == 7){
				my @infoList = split(';', $line[$i] );	
				foreach my $info (@infoList){
					my @ii = split('=', $info );
					if (scalar @ii == 2){
						if ($ii[0] eq "GESP") {
							my @GESP_GT = split('/', $ii[1] );
							if (scalar @GESP_GT == 2) {
								my $nbAlt = 0;
								my $nbRef = 0;
								
								if ($GESP_GT[0] =~ /$alt/){
									$nbAlt = (split ':',$GESP_GT[0])[1];
									#print "alt: ",$alt,"catch: ",$GESP_GT[0],"value: ",$nbAlt,"\n";		
								}	
								if ($GESP_GT[1] =~ /$alt/){
									$nbAlt = (split ':',$GESP_GT[1])[1];
								}	
								if ($GESP_GT[0] =~ /$ref/){
									$nbRef = (split ':',$GESP_GT[0])[1];
								}	
								if ($GESP_GT[1] =~ /$ref/){
									$nbRef = (split ':',$GESP_GT[1])[1];
								}	
								if ($nbRef > 0 && $nbAlt > 0 ){
									$freqGESP = $nbAlt / ($nbRef + $nbAlt);
								}	


							}
						}


						if($caller eq "freebayes"){

							if ($ii[0] eq "GEXAC" && $alt !~ /,/  ) {

								#DEBUG
								#print $line[7],"\t";

								if ($ii[1] =~ /ref/){

									if (length $alt > length $ref){
										$alt = "ins".substr($alt,(length $ref)-1,((length $alt)-(length $ref)));
										
										#DEBUG print "alt ",$alt,"\n";
										
									}elsif(length $alt < length $ref){
										$alt = "del".substr($ref,(length $alt)-1,((length $ref)-(length $alt)));
										#DEBUG print "del ",$alt,"\n";
									}elsif(length $alt == length $ref){
										$alt = "mnp";
										
									}
									$ref="ref";
								}
								
								my @GEXAC_GT = split('/', $ii[1] );
								if (scalar @GEXAC_GT > 1) {
								
										#print $ii[1],"\n";
								
									my $nbAltx = 0;
									#my $nbRefx = 0;
									my $alleleNumber = 0;

									foreach my $var (@GEXAC_GT){
										$alleleNumber += (split ':', $var)[1];
									
										if ((split ':', $var)[0] eq $alt){
											$nbAltx = (split ':', $var)[1];

										}	
									}
								
								#print $alleleNumber,"\t",$nbAltx,"\n";
								
									if ($alleleNumber > 0 && $nbAltx > 0 ) {
										$freqEXAC = $nbAltx / $alleleNumber;
									}
								}
							}
						}else {
							if ($ii[0] eq "GEXAC" && $alt !~ /,/  ) {
								if ($ii[1] =~ /ref/){
									if (length $alt > length $ref){
										$alt = "ins".substr($alt,1);
										#print $alt,"\t";
									}else{
										$alt = "del".substr($ref,1);
										#print "ici:   ".$alt."\t".substr($alt,1);
									}
									$ref="ref";
								}
								my @GEXAC_GT = split('/', $ii[1] );
								if (scalar @GEXAC_GT > 1) {
								
										#print $ii[1],"\n";
								
									my $nbAltx = 0;
									my $nbRefx = 0;
									my $alleleNumber = 0;

									foreach my $var (@GEXAC_GT){
										$alleleNumber += (split ':', $var)[1];
										#print "split:  ".(split ':', $var)[0]."\t\t".$alt."\n";	
										if ((split ':', $var)[0] eq $alt){
											$nbAltx = (split ':', $var)[1];

										}	
									}
								
								#print $alleleNumber,"\t",$nbAltx,"\n";
								
									if ($alleleNumber > 0 && $nbAltx > 0 ) {
										$freqEXAC = $nbAltx / $alleleNumber;
									}
								}
							}
							
							
							
							
						}
						
						
						
						#check mutation function ns/ss/i
						if ($ii[0] eq "FG"){
							my @fg = split ',', $ii[1];
							my $sommeFG =0;
							foreach my $f (@fg){
								if (defined $dicoFonction{$f}){
                  #$dicoFonction{$f} += 1;
									$sommeFG ++;
								} 
							}
							if ($sommeFG > 0){
								$nsssi = "true";
							} else {
								$nsssi = "false";
							}
						} 
			#print "Param:\t",$ii[0],"\tet\t",$ii[1],"\n";
						my ($index) = grep {$colonnes[$_] eq $dicoParam{$ii[0]} } (0 .. @colonnes-1);

						if (defined $index){
							$data{$index} = $ii[1];

						}else{
							print "Problème: le paramètre ",$ii[0]," n'est pas présent dans la liste des colonnes. ".$colonnes[$_]. "\n";
							exit 0;
						}
						#print "index:\t",defined $index ? $index : -1 , "\n";

						#print "test\t", first {$colonnes[$_] eq $ii[0] } 0..$#colonnes , "\n";
					}
				}
			}

			#Parse GEnotypes
			if ($i > 8){
				my @genotype = split(':', $line[$i] );
				
								
				$data{$i-2+$nbrSample} = substr $line[$i], 0, 3;


				my $DP;
				my $adalt;
				my $adref;
				my $AB;
				my $AD;
				
				#if (scalar @genotype > 1 && $line[8] =~ /:AD:/ ){
				
				if (scalar @genotype > 1 && $caller eq "freebayes"){
					#$AD = $genotype[3].",".$genotype[5];
					
					#DEBUG
					#print $genotype[2];

					if ($genotype[0] eq "."){
						$DP = 0;
						$AD = "0,0";
					}else {
						$DP = $genotype[3] + $genotype[5];
						$AD = $genotype[3].",".$genotype[5];
					}
					
					my @tabAD = split( ',',$AD);
					
					#DEBUG
					#print "\ntabAD\t",split( ',',$AD),"\n";

					if ( scalar @tabAD > 1 ){
						$adref = $tabAD[0];
						my $totalAD = $adref;
							
					

						for (my $j = 1; $j < scalar @tabAD; $j++){
							$totalAD += $tabAD[$j];
						}

						if ($totalAD == 0){
							$AB = 0;
						}else{
							for (my $j = 1; $j < scalar @tabAD; $j++){
								
								$AB .= substr(($tabAD[$j]/ $totalAD),0,5);
								$AB .= ",";

							}


							$AB= substr $AB, 0 , ((length $AB)-1);


						}

					}
				
				}elsif (scalar @genotype > 1 && $caller eq "GATK"){


						if ($genotype[2] eq "."){
							$DP = 0;
							$AD = "0,0";

						}else {
							$DP = $genotype[2];
							$AD = $genotype[1];
						
						}
					
					#DEBUG
					#print $genotype[2]."\n";


					
					my @tabAD = split( ',',$AD);
					
					#print "\ntabAD\t",split( ',',$AD),"\n";

					if ( scalar @tabAD > 1 ){
						$adref = $tabAD[0];
						my $totalAD = $adref;
							
					

						for (my $j = 1; $j < scalar @tabAD; $j++){
							$totalAD += $tabAD[$j];
						}

						if ($totalAD == 0){
							$AB = 0;
						}else{
							for (my $j = 1; $j < scalar @tabAD; $j++){
								
								$AB .= substr(($tabAD[$j]/ $totalAD),0,5);
								$AB .= ",";

							}


							$AB= substr $AB, 0 , ((length $AB)-1);


						}

					}else {
						$adref = 0;
						$adalt= 0;
						$AB = 0;
					}

				}else {
					$adref = 0;
					$adalt= 0;
					$AB = 0;
					$DP = 0;
					$AD = "0,0";
				}


				$data{$i-2} = "|GT=".$genotype[0]."|AD=".$AD."|AB=".$AB."|DP=".$DP."|"; 

			}

		}

		for(my $i = 0; $i < scalar @colonnes; $i ++){
			if(defined $data{$i}){
				$chaine .= $data{$i}."\t";
			}else{
				$chaine .= "\t";
			}
		}
    



		if($freqGESP > 0){
			$chaine .= ($freqGESP*100)."\t";
		} else{
			$chaine .= "0\t";
		}
		
		$chaine .= $nsssi."\t";
		$chaine .= "http://exac.broadinstitute.org/region/".$data{0}."-".$data{1}."-".$data{1}."\t";

		if($freqEXAC > 0){
			$chaine .= ($freqEXAC*100)."\t";
		} else{
			$chaine .= "0\t";
		}

		$chaine .= "\n";
		print OUT $chaine;


		chomp $chaine;

		#sort final chaine and print
    		@unorderedLine = split( /\t/, $chaine );

		foreach my $finaleCol (sort {$a <=> $b}  keys %dicoOrderColumns) {

			#print $unorderedLine[$dicoOrderColumns{$finalCol}{'colIndex'}]; 
			$orderedChaine .= $unorderedLine[$dicoOrderColumns{$finaleCol}{'colIndex'}]."\t";
#			print $orderedChaine."\n". $finaleCol; 
		}
		
		$orderedChaine .= "\n";

#		print OUTSHORTSORT $orderedChaine;

		@orderedLine = split( /\t/,$orderedChaine);
	

		#SELECT only lines with  PASS FILTER, and index = ./. and parents != ./. 
		#Find deletion homozygote regions
		if ($orderedLine[8+($cmpt/2) ] eq "./." && $orderedLine[9+($cmpt/2) ] ne "./." && $orderedLine[10+($cmpt/2) ] ne "./." && $orderedLine[6] eq "PASS" && $trio eq "YES" && ($orderedLine[10+$cmpt] ne "")  ){
			$worksheetDELHMZ->write_row( $worksheetLineDELHMZ, 0, \@orderedLine );
			#$worksheetDELHMZ->write( $worksheetLineDELHMZ, 9+$cmpt,$geneSummaryConcat,$format_pLI );
			#if($pLI_values ne ""){
			#	$worksheetDELHMZ->write_comment( $worksheetLineDELHMZ, 9+$cmpt,$pLI_values,x_scale => 2 );
			#}
			$worksheetLineDELHMZ ++;
		}







		#SELECT only lines with  PASS FILTER, ExaC freq < 1% and nsssi = true
#		if (($orderedLine[6] eq "PASS") && ($orderedLine[18+$cmpt] lt 1) && ($orderedLine[34+$cmpt] eq "true") ){
		if (($orderedLine[6] eq "PASS") && ($orderedLine[18+$cmpt] lt 1) && ($orderedLine[33+$cmpt] eq "true") ){
#			print OUTPASStrue1pct $orderedChaine;


			if(defined $dicoGeneSummary{$orderedLine[9+$cmpt]}){
				$geneSummaryConcat = $orderedLine[9+$cmpt]." ### ".$dicoGeneSummary{$orderedLine[9+$cmpt]};
			}else{
		        	$geneSummaryConcat = $orderedLine[9+$cmpt];
			}



			if(defined $dicopLI{$orderedLine[9+$cmpt]}{"comment"}){
				$pLI_values = $dicopLI{$orderedLine[9+$cmpt]}{"comment"};
				#$format_pLI = sprintf("#%2.2X%2.2X%2.2X\n",($dicopLI{$orderedLine[9+$cmpt]}{"pLI"}*255 + $dicopLI{$orderedLine[9+$cmpt]}{"pRec"}*255),($dicopLI{$orderedLine[9+$cmpt]}{"pRec"}*255 + $dicopLI{$orderedLine[9+$cmpt]}{"pNull"} * 255),0);
#				$format_pLI -> set_bg_color($dicopLI{$orderedLine[9+$cmpt]}{"color_format"});
				#print $dicopLI{$orderedLine[9+$cmpt]}{"color_format"}."\t".$format_pLI."\n";
 
        			$format_pLI = $workbook->add_format(bg_color => $dicopLI{$orderedLine[9+$cmpt]}{"color_format"});

			}else{
		        	$pLI_values = "";
              			#$format_pLI -> set_bg_color('#FFFFFF');
				$format_pLI = $workbook->add_format(bg_color => '#FFFFFF');

			} 



			if($orderedLine[10+$cmpt] ne ""){
      
				$worksheetOMIM->write_row( $worksheetLineOMIM, 0, \@orderedLine );
				$worksheetOMIM->write( $worksheetLineOMIM, 9+$cmpt,$geneSummaryConcat,$format_pLI );
				if($pLI_values ne ""){
					$worksheetOMIM->write_comment( $worksheetLineOMIM, 9+$cmpt,$pLI_values ,x_scale => 2);
				}

				#print $orderedLine[10+$cmpt]."\n";       
				$worksheetLineOMIM ++;
      			}

			
			
			foreach my $ACMGgene (@ACMGlist) {
				if($orderedLine[9+$cmpt] eq $ACMGgene){
					$worksheetACMG->write_row( $worksheetLineACMG, 0, \@orderedLine );
					$worksheetACMG->write( $worksheetLineACMG, 9+$cmpt,$geneSummaryConcat,$format_pLI );
					if($pLI_values ne ""){
						$worksheetACMG->write_comment( $worksheetLineACMG, 9+$cmpt,$pLI_values,x_scale => 2 );
					}	
					$worksheetLineACMG ++;
				}
			}
			
	    		if($candidates ne ""){
				foreach my $candGene (@candidatesList){
					if($orderedLine[9+$cmpt] eq $candGene){
						$worksheetCandidats->write_row( $worksheetLineCandidats, 0, \@orderedLine );
						$worksheetCandidats->write( $worksheetLineCandidats, 9+$cmpt,$geneSummaryConcat,$format_pLI );
						if($pLI_values ne ""){
							$worksheetCandidats->write_comment( $worksheetLineCandidats, 9+$cmpt,$pLI_values ,x_scale => 2);
						}	
						$worksheetLineCandidats ++;
					}
				}
			}

			#additionnal analysis in TRIO context
			if ($trio eq "YES"){

				if($orderedLine[8+($cmpt/2) ] eq "0/1" && (($orderedLine[9+($cmpt/2) ] eq "0/0" && ($orderedLine[10+($cmpt/2) ] eq "0/1" ||$orderedLine[10+($cmpt/2) ] eq "1/0" ))|| (($orderedLine[9+($cmpt/2) ] eq "0/1" ||  $orderedLine[9+($cmpt/2) ] eq "1/0" ) && $orderedLine[10+($cmpt/2) ] eq "0/0"  ))){
					
					#Find HTZ composite
					if(defined $dicoGeneForHTZcompo{$orderedLine[9+$cmpt]} ){
						
						$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'cnt'} ++;

						if($orderedLine[9+($cmpt/2) ] eq "0/0"){

							$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'pvsm'} ++;

							if( $dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'mvsp'} >= 1){  
								$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'ok'} = 1;
							}
						}else{
							$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'mvsp'} ++;

							if( $dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'pvsm'} >= 1){  
								$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'ok'} = 1;
							}
						}

						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@emptyArray );
						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@orderedLine );
						$worksheetHTZcompo->write( $worksheetLineHTZcompo, 9+$cmpt,$geneSummaryConcat,$format_pLI );
						if($pLI_values ne ""){
							$worksheetHTZcompo->write_comment( $worksheetLineHTZcompo, 9+$cmpt,$pLI_values,x_scale => 2);
						}
						$worksheetLineHTZcompo ++;

					}else{

#						if(defined $dicoGeneForHTZcompo{$previousGene} && $dicoGeneForHTZcompo{$previousGene}{'ok'}==0 ){
						if(($previousGene ne $orderedLine[9+$cmpt]) && $dicoGeneForHTZcompo{$previousGene}{'ok'}==0 ){
								$worksheetLineHTZcompo -= $dicoGeneForHTZcompo{$previousGene}{'cnt'};

						}
							
						$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'cnt'} = 1;
						$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'ok'} = 0;

						if($orderedLine[9+($cmpt/2) ] eq "0/0"){

							$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'pvsm'} = 1;
							$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'mvsp'} = 0;
						
						}else{
							$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'mvsp'} = 1;
							$dicoGeneForHTZcompo{$orderedLine[9+$cmpt]}{'pvsm'} = 0;
						}

						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@emptyArray );
						$worksheetHTZcompo->write_comment( $worksheetLineHTZcompo, 9+$cmpt,"",x_scale => 2);
						
						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@orderedLine );
						$worksheetHTZcompo->write( $worksheetLineHTZcompo, 9+$cmpt,$geneSummaryConcat,$format_pLI );
						if($pLI_values ne ""){
							$worksheetHTZcompo->write_comment( $worksheetLineHTZcompo, 9+$cmpt,$pLI_values,x_scale => 2 );
						}
						$worksheetLineHTZcompo ++;
						#print $orderedLine[9+$cmpt]."\n";
						
						$previousGene = $orderedLine[9+$cmpt];
					}


				}
				
				#Find Homozygote recessive AR
				if($orderedLine[8+($cmpt/2) ] eq "1/1" && ( $orderedLine[9+($cmpt/2) ] eq "0/1" || $orderedLine[9+($cmpt/2) ] eq "1/0"   ) && ( $orderedLine[10+($cmpt/2) ] eq "0/1" || $orderedLine[10+($cmpt/2) ] eq "1/0"  ) ){
					$worksheetAR->write_row($worksheetLineAR , 0, \@orderedLine );
					$worksheetAR->write( $worksheetLineAR, 9+$cmpt,$geneSummaryConcat,$format_pLI );
					if($pLI_values ne ""){
						$worksheetAR->write_comment( $worksheetLineAR, 9+$cmpt,$pLI_values ,x_scale => 2);
					}	
					$worksheetLineAR ++; 
				}
				
				#Find SNVvsCNV
				if($orderedLine[8+($cmpt/2) ] eq "1/1" && $orderedLine[9+($cmpt/2) ] eq "0/0"  && ( $orderedLine[10+($cmpt/2) ] eq "0/1" || $orderedLine[10+($cmpt/2) ] eq "1/0")){
					$worksheetSNPmereVsCNVpere->write_row($worksheetLineSNPmereVsCNVpere , 0, \@orderedLine );
					$worksheetSNPmereVsCNVpere->write( $worksheetLineSNPmereVsCNVpere, 9+$cmpt,$geneSummaryConcat,$format_pLI );
					if($pLI_values ne ""){
						$worksheetSNPmereVsCNVpere->write_comment( $worksheetLineSNPmereVsCNVpere, 9+$cmpt,$pLI_values,x_scale => 2 );
					}
					$worksheetLineSNPmereVsCNVpere ++;
				}

				if($orderedLine[8+($cmpt/2) ] eq "1/1" && ( $orderedLine[9+($cmpt/2) ] eq "0/1"  || $orderedLine[9+($cmpt/2) ] eq "1/0"   ) && $orderedLine[10+($cmpt/2) ] eq "0/0"){
					$worksheetSNPpereVsCNVmere->write_row($worksheetLineSNPpereVsCNVmere , 0, \@orderedLine );
					$worksheetSNPpereVsCNVmere->write( $worksheetLineSNPpereVsCNVmere, 9+$cmpt,$geneSummaryConcat,$format_pLI );
					if($pLI_values ne ""){
						$worksheetSNPpereVsCNVmere->write_comment( $worksheetLineSNPpereVsCNVmere, 9+$cmpt,$pLI_values ,x_scale => 2);
					}
					$worksheetLineSNPpereVsCNVmere ++;
				}

				#Find de novo
				if(($orderedLine[8+($cmpt/2) ] eq "1/1" || $orderedLine[8+($cmpt/2) ] eq "0/1" || $orderedLine[8+($cmpt/2) ] eq "1/0"  ) && $orderedLine[9+($cmpt/2) ] eq "0/0" && $orderedLine[10+($cmpt/2) ] eq "0/0"){
					$worksheetDENOVO->write_row( $worksheetLineDENOVO, 0, \@orderedLine );
					$worksheetDENOVO->write( $worksheetLineDENOVO, 9+$cmpt,$geneSummaryConcat,$format_pLI );
					if($pLI_values ne ""){
						$worksheetDENOVO->write_comment( $worksheetLineDENOVO, 9+$cmpt,$pLI_values,x_scale => 2 );
					}
					$worksheetLineDENOVO ++;
				}

				


			}
			
      #write sheet "ALL"
			$worksheet->write_row( $worksheetLine, 0, \@orderedLine );

			$worksheet->write( $worksheetLine, 9+$cmpt,$geneSummaryConcat,$format_pLI );
			
			#      print $orderedLine[9+$cmpt]."\t".$dicopLI{$orderedLine[9+$cmpt]}{"color_format"}."\t".$format_pLI."\n";
			
			if($pLI_values ne ""){
				$worksheet->write_comment( $worksheetLine, 9+$cmpt,$pLI_values ,x_scale => 2);
			}
			$worksheetLine ++;

		}
		
		

	}#END of SELECT only lines with  PASS FILTER, ExaC freq < 1% and nsssi = true

	
	#SELECT only lines with NON-PASS FILTER, ExaC freq < 1% and nsssi = true and OMIM != ""
	if (($orderedLine[6] ne "PASS") && ($orderedLine[18+$cmpt] lt 1) && ($orderedLine[33+$cmpt] eq "true") && ($orderedLine[10+$cmpt] ne "")   ){
		#print in another sheet excel or file
		$worksheetRAREnonPASS->write_row( $worksheetLineRAREnonPASS, 0, \@orderedLine );
		#$worksheetDELHMZ->write( $worksheetLineDELHMZ, 9+$cmpt,$geneSummaryConcat,$format_pLI );
		#if($pLI_values ne ""){
		#	$worksheetDELHMZ->write_comment( $worksheetLineDELHMZ, 9+$cmpt,$pLI_values,x_scale => 2 );
		#}
		$worksheetLineRAREnonPASS ++;

	}

	#SELECT only lines with CLinVar Pathogenic , ExaC freq < 1%
	if (($orderedLine[31+$cmpt] =~ /athogeni/) &&  ($orderedLine[31+$cmpt] =~ /^[LP]/)  && ($orderedLine[18+$cmpt] lt 1)){
		#print in another sheet excel or file
		$worksheetClinVarPatho->write_row( $worksheetLineClinVarPatho, 0, \@orderedLine );
		#$worksheetDELHMZ->write( $worksheetLineDELHMZ, 9+$cmpt,$geneSummaryConcat,$format_pLI );
		#if($pLI_values ne ""){
		#	$worksheetDELHMZ->write_comment( $worksheetLineDELHMZ, 9+$cmpt,$pLI_values,x_scale => 2 );
		#}
		$worksheetLineClinVarPatho ++;

	}


}



close(VCF);
close(OUT);
#close(OUTSHORTSORT);
print STDERR "Done!";

exit 0;

