#!/usr/bin/perl

##### achab.pl ####

# Author : Thomas Guignard 2018

# Description : 
# Create an User friendly Excel file from an MPA annotated VCF file. 


use strict; 
use warnings;
use Getopt::Long; 
use Excel::Writer::XLSX;
use Switch;
#use Pod::Usage;
#use List::Util qw(first);
#use Data::Dumper;

#parameters
my $man = "USAGE : \nperl achab.pl 
\n--vcf <vcf_file> 
\n--outDir <output directory (default = current dir)> 
\n--candidates <file with gene symbol of interest>  
\n--phenolyzerFile <phenolyzer output file suffixed by predicted_gene_scores>   
\n--popFreqThr <allelic frequency threshold from 0 to 1 default=0.01> 
\n--trio (requires case dad and mum option to be filled) 
\n\t--case <index_sample_name> 
\n\t--dad <father_sample_name> 
\n\t--mum <mother_sample_name>  
\n--customInfoList  <comma separated list of vcf-info names (will be added in a new column)>  
\n--filterList <comma separated list of VCF FILTER to output (default='PASS', included )>   
\n--cnvGeneList <File with gene symbol + annotation , involved by parallel CNV calling >
\n--newHope (output only NON PASS or MPA_rank = 8 variants, default=output FILTER=PASS and MPAranking < 8 variants )>";

my $help;
my $current_line;
my $incfile = "";
my $outDir = ".";
my $case = "";
my $mum = "";
my $dad = "";
my $caller = "";
my $trio;
my $popFreqThr = "";
my $filterList = "";
my @filterArray;
my $newHope;
my $sampleNames = "";
my @sampleList;
my $customInfoList = "";
my @custInfList;

#stuff for files
my $candidates = "";
my %candidateGene;
my $candidates_line;

my $phenolyzerFile = "";
my $phenolyzer_Line;
my @phenolyzer_List;
my %phenolyzerGene;

my $cnvGeneList = "";
my %cnvGene;
my $cnvGene_Line;
my @cnvGene_List;


#vcf parsing and output
my @line;
my $variantID; 

my @finalSortData;
my $familyGenotype;		
my %hashFinalSortData;
	

#$arguments = GetOptions( "vcf=s" => \$incfile ) or pod2usage(-vcf => "$0: argument required\n") ;

GetOptions( 	"vcf=s"				=> \$incfile,
		"case=s"			=> \$case,
		"dad=s"				=> \$dad, 
		"mum=s"				=> \$mum, 
		"trio"				=> \$trio,
		"candidates:s"			=> \$candidates,
		"outDir=s"			=> \$outDir,
		"phenolyzerFile=s"		=> \$phenolyzerFile,
		"popFreqThr=s"			=> \$popFreqThr, 
		"customInfoList:s"			=> \$customInfoList, 
		"filterList:s"			=> \$filterList,				
		"cnvGeneList:s"			=> \$cnvGeneList,
		"newHope"			=> \$newHope,
		"help|h"				=> \$help);
				

#check mandatory arguments
if(defined $help || $incfile eq ""){
	die("$man");
}

#define popFreqThr
if( $popFreqThr eq ""){
	$popFreqThr = 0.01;
	
}

#define filter List
if($filterList ne ""){
	@filterArray = split(/,/ , $filterList)
}
#default filter is PASS
unshift @filterArray, "PASS";



#check sample list param
if(defined $trio && ($case eq "" || $dad eq "" || $mum eq "")){	
	die("TRIO option requires 3 sample names. Please, give --case, --dad and --mum sample name arguments.\n");
}
			
			
print  STDERR "Starting a new fishing trip ... \n" ; 
print  STDERR "Processing vcf file ... \n" ; 


open( VCF , "<$incfile" )or die("Cannot open vcf file $incfile") ;


#TODO check if header contains required INFO
#Parse VCF header to fill the dictionnary of parameters
print STDERR "Parsing VCF header in order to get sample names and to check if required informations are present ... \n";
my %dicoParam;

while( <VCF> ){
  	$current_line = $_;
		
	chomp $current_line;
      
    #filling dicoParam with VCF header INFO and FORMAT 

    if ($current_line=~/^##/){


		  unless ($current_line=~/Description=/){ next }
      #DEBUG print STDERR "Header line\n";

      if ($current_line =~ /ID=(.+?),.*?Description="(.+?)"/){
    
          $dicoParam{$1}= $2;
		  #DEBUG      print STDERR "info : ". $1 . "\tdescription: ". $2."\n";
			

			    next;
      
      }else {print STDERR "pattern not found in this line: ".$current_line ."\n";next} 
			
	}elsif($current_line=~/^#CHROM/){
		#check sample names or die
	
		
		if (defined $trio){
			#check if case sample name is found
			unless ($current_line=~/\Q$case/){die("$case is not found as a sample in the VCF, please check case name.\n$current_line\n")}
			unless ($current_line=~/\Q$dad/){die("$dad is not found as a sample in the VCF, please check dad name.\n$current_line\n")}
			unless ($current_line=~/\Q$mum/){die("$mum is not found as a sample in the VCF, please check mum name.\n$current_line\n")}
		}

		@line = split (/\t/ , $current_line);


		for( my $sampleIndex = 9 ; $sampleIndex < scalar @line; $sampleIndex++){
			
			print STDERR "Found Sample ".$line[$sampleIndex]."\n";
			push @sampleList, $line[$sampleIndex]; 
			
			#for each final position for sample 
#			foreach my $finCol (keys %dicoSamples){

    	    #DEBUG
					
#				if ($dicoSamples{$finCol}{'columnName'} eq "Genotype-".$line[$sampleIndex] ){
#					$dicoSamples{$finCol}{'columnIndex'} =  $sampleIndex;				
#					last;
#				}
#			}

#			foreach my $name (@sampleList)
				#if($line[$sampleIndex]
				
			
		}	

		if($case eq ""){
			$case = $sampleList[0];
		}
		#exclude to treat trio with too much sample
		print STDERR "\nTotal Samples : ".scalar @sampleList."\n";

		if(scalar @sampleList > 3 && defined $trio){
			die("Found more than 3 samples. TRIO analysis is not supported with more than 3 samples.\n");
		}



	}else {last}

}
close(VCF);



# Create a new Excel workbook
#
my $workbook;

if ($outDir eq "." || -d $outDir) {
	if(defined $newHope){
		# Create a "new hope Excel" aka NON-PASS + MPA_RANKING=8 variants 
		$workbook = Excel::Writer::XLSX->new( $outDir."/achab_catch_newHope.xlsx" );
	}else{
		$workbook = Excel::Writer::XLSX->new( $outDir."/achab_catch.xlsx" );
	}
}else {
 	die("No directory $outDir");
}


#create default color background when pLI values are absent
my $format_pLI = $workbook->add_format(bg_color => '#FFFFFF');
#$format_pLI -> set_pattern();


# Add all worksheets
my $worksheet = $workbook->add_worksheet('ALL_'.$popFreqThr);
$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLine = 0;

my $worksheetACMG = $workbook->add_worksheet('DS_ACMG');
$worksheetACMG->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineACMG = 0;





#Worksheets initialization

my $worksheetHTZcompo;
my $worksheetAR; 
my $worksheetSNPmumVsCNVdad ;
my $worksheetSNPdadVsCNVmum;
my $worksheetDENOVO;
my $worksheetCandidats;
#my $worksheetDELHMZ;

my $worksheetLineHTZcompo ;
my $worksheetLineAR ; 
my $worksheetLineSNPmumVsCNVdad ;
my $worksheetLineSNPdadVsCNVmum ;
my $worksheetLineDENOVO ;
my $worksheetLineCandidats;
#my $worksheetLineDELHMZ;


#create additionnal sheet in trio analysis
if (defined $trio){
	$worksheetHTZcompo = $workbook->add_worksheet('HTZ_compo');
	$worksheetLineHTZcompo = 0;
	$worksheetHTZcompo->freeze_panes( 1, 0 );    # Freeze the first row
  
	$worksheetAR = $workbook->add_worksheet('AR');
	$worksheetLineAR = 0;
	$worksheetAR->freeze_panes( 1, 0 );    # Freeze the first row
  
	$worksheetSNPmumVsCNVdad = $workbook->add_worksheet('SNVmumVsCNVdad');
	$worksheetLineSNPmumVsCNVdad = 0;
	$worksheetSNPmumVsCNVdad->freeze_panes( 1, 0 );    # Freeze the first row
  
	$worksheetSNPdadVsCNVmum = $workbook->add_worksheet('SNVdadVsCNVmum');
	$worksheetLineSNPdadVsCNVmum = 0;
	$worksheetSNPdadVsCNVmum->freeze_panes( 1, 0 );    # Freeze the first row
  
	$worksheetDENOVO = $workbook->add_worksheet('DENOVO');
	$worksheetLineDENOVO = 0;
	$worksheetDENOVO->freeze_panes( 1, 0 );    # Freeze the first row

	#$worksheetDELHMZ = $workbook->add_worksheet('DEL_HMZ');
	#$worksheetLineDELHMZ = 0;
	#$worksheetDELHMZ->freeze_panes( 1, 0 );    # Freeze the first row
	#$worksheetDELHMZ->autofilter('A1:AN1000'); # Add autofilter
		

}




#get data from phenolyzer output (predicted_gene_score)
my $current_gene= "";
my $maxLine=0;
if($phenolyzerFile ne ""){
	open(PHENO , "<$phenolyzerFile") or die("Cannot open phenolyzer file ".$phenolyzerFile) ;
	print  STDERR "Processing phenolyzer file ... \n" ; 
	while( <PHENO> ){
		next if($_ =~/^Tuple number/);
	
		$phenolyzer_Line = $_;
		chomp $phenolyzer_Line;
		
		if($phenolyzer_Line=~/ID:/){
			
			#should write it for previous gene, not the current one $phenolyzerGene{$current_gene}{'comment'} = $maxLine." lines of features.\n";

			@phenolyzer_List = split( /\t/, $phenolyzer_Line);
			$current_gene = $phenolyzer_List[0];
			$phenolyzerGene{$current_gene}{'Raw'}= $phenolyzer_List[2];
			$phenolyzerGene{$current_gene}{'comment'}= $phenolyzer_List[1]."\n".$phenolyzer_List[3]."\n";
			$maxLine=0;

		}else{	
			#keep only OMIM or count nbr of lines
			$maxLine+=1;
					
		} 			
	}
	close(PHENO);
} 



#create sheet for candidats
if($candidates ne ""){
	open( CANDIDATS , "<$candidates")or die("Cannot open candidates file ".$candidates) ;
	print  STDERR "Processing candidates file ... \n" ; 
	while( <CANDIDATS> ){
	  	$candidates_line = $_;
		chomp $candidates_line;
		$candidateGene{$candidates_line} = 1;

				
	}
	close(CANDIDATS);
	$worksheetCandidats = $workbook->add_worksheet('Candidats');
	$worksheetCandidats->freeze_panes( 1, 0 );    # Freeze the first row
}


#get gene involved in CNV
if ($cnvGeneList ne ""){
	open( CNVGENES , "<$cnvGeneList")or die("Cannot open cnvGeneList file ".$cnvGeneList) ;
	print  STDERR "Processing CNV Gene file ... \n" ; 
	while( <CNVGENES> ){
	  	$cnvGene_Line = $_;
		chomp $cnvGene_Line;
		@cnvGene_List = split( /\t/, $cnvGene_Line);
		$current_gene = $cnvGene_List[0];
		if (defined $cnvGene{$current_gene}){
				#nothing to do
		}else{	
			$cnvGene{$current_gene} = "CNV : ".$current_gene;
		}

		if(defined $cnvGene_List[1]){
				$cnvGene{$current_gene} .= " \n=> ".$cnvGene_List[1];
		}
			
	}
	close(CNVGENES);

}




#Hash of ACMG incidentalome genes
my %ACMGgene = ("ACTA2" =>1,"ACTC1" =>1,"APC" =>1,"APOB" =>1,"ATP7B" =>1,"BMPR1A" =>1,"BRCA1" =>1,"BRCA2" =>1,"CACNA1S" =>1,"COL3A1" =>1,"DSC2" =>1,"DSG2" =>1,"DSP" =>1,"FBN1" =>1,"GLA" =>1,"KCNH2" =>1,"KCNQ1" =>1,"LDLR" =>1,"LMNA" =>1,"MEN1" =>1,"MLH1" =>1,"MSH2" =>1,"MSH6" =>1,"MUTYH" =>1,"MYBPC3" =>1,"MYH11" =>1,"MYH7" =>1,"MYL2" =>1,"MYL3" =>1,"NF2" =>1,"OTC" =>1,"PCSK9" =>1,"PKP2" =>1,"PMS2" =>1,"PRKAG2" =>1,"PTEN" =>1,"RB1" =>1,"RET" =>1,"RYR1" =>1,"RYR2" =>1,"SCN5A" =>1,"SDHAF2" =>1,"SDHB" =>1,"SDHC" =>1,"SDHD" =>1,"SMAD3" =>1,"SMAD4" =>1,"STK11" =>1,"TGFBR1" =>1,"TGFBR2" =>1,"TMEM43" =>1,"TNNI3" =>1,"TNNT2" =>1,"TP53" =>1,"TPM1" =>1,"TSC1" =>1,"TSC2" =>1,"VHL" =>1,"WT1"=>1);



                  
#empty line to erase false HTZ composite lines
my @emptyArray = (" "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "," "); 




 


#counter for shifting columns according to nbr of sample

my $cmpt = scalar @sampleList;


#dico for sample sorting (index / dad / mum / control)
my %dicoSamples ;


#
my %dicoColumnNbr;
$dicoColumnNbr{'MPA_ranking'}=				0;	#+ comment MPA scores and related scores 
$dicoColumnNbr{'Phenolyzer'}=				1;	#Phenolyzer raw score + comment Normalized score 
$dicoColumnNbr{'Gene.refGene'}=				2;  #Gene Name + comment pLi / Function_description / tissue specificity
$dicoColumnNbr{'Phenotypes.refGene'}=			3;  #OMIM + comment Disease_description
$dicoColumnNbr{'gnomAD_genome_ALL'}=			4;	#Pop Freq + comment ethny
$dicoColumnNbr{'gnomAD_exome_ALL'}=			5;	#as well
$dicoColumnNbr{'CLINSIG'}=				6;	#CLinvar
$dicoColumnNbr{'InterVar_automated'}=			7;	#+ comment ACMG status
$dicoColumnNbr{'SecondHit-CNV'}=			8;	#TODO
$dicoColumnNbr{'Func.refGene'}=				9;	# + comment ExonicFunc / AAChange / GeneDetail



if (defined $trio){
	
		$dicoColumnNbr{'Genotype-'.$case}=		10;
		$dicoColumnNbr{'Genotype-'.$dad}=		11;
		$dicoColumnNbr{'Genotype-'.$mum}=		12;

		$dicoSamples{1}{'columnName'} = 'Genotype-'.$case ;
		$dicoSamples{2}{'columnName'} = 'Genotype-'.$dad ;
		$dicoSamples{3}{'columnName'} = 'Genotype-'.$mum ;

	
}else{

	for( my $i = 0 ; $i < scalar @sampleList; $i++){
	
		$dicoColumnNbr{'Genotype-'.$sampleList[$i]}=			10+$i;	# + comment qual / caller / DP AD AB
		$dicoSamples{$i+1}{'columnName'} = 'Genotype-'.$sampleList[$i] ;
	}

}


$dicoColumnNbr{'#CHROM'}=	10+$cmpt ;
$dicoColumnNbr{'POS'}=		11+$cmpt ;
$dicoColumnNbr{'ID'}=		12+$cmpt ;
$dicoColumnNbr{'REF'}=		13+$cmpt ;
$dicoColumnNbr{'ALT'}=		14+$cmpt ;
$dicoColumnNbr{'FILTER'}=	15+$cmpt ;

#Add custom Info in additionnal columns
my $lastColumn = 16+$cmpt; 

if($customInfoList ne ""){
	@custInfList = split(/,/, $customInfoList);
	foreach my $customInfo (@custInfList){
		$dicoColumnNbr{$customInfo} = $lastColumn ;
		$lastColumn++;
	}
}



#Define column title order
my @columnTitles;
foreach my $key  (sort { $dicoColumnNbr{$a} <=> $dicoColumnNbr{$b} } keys %dicoColumnNbr)  {
	push @columnTitles,  $key;
	#DEBUG print STDERR $key."\n";
}


#final strings for comment
my $commentGenotype;
my $commentMPAscore;
my $commentGnomADExomeScore;
my $commentGnomADGenomeScore;
my $commentInterVar;
my $commentFunc;
my $commentPhenotype;

#define sorted arrays with score for comment
my @CommentMPA_score = ("MPA_final_score",
						"MPA_impact",
						"MPA_adjusted",
						"MPA_available",
						"MPA_deleterious",
						"\n--- SPLICE ---",
						'dbscSNV_ADA_SCORE',
						'dbscSNV_RF_SCORE',
						'dpsi_zscore',
						"\n--- MISSENSE ---",
						'LRT_pred',
						'SIFT_pred',
						'FATHMM_pred',
						'MetaLR_pred',
						'MetaSVM_pred',
						'PROVEAN_pred',
						'Polyphen2_HDIV_pred',
						'Polyphen2_HVAR_pred',
						'MutationTaster_pred',
						'fathmm-MKL_coding_pred');

my $pLI_Comment = "pLI - the probability of being loss-of-function intolerant (intolerant of both heterozygous and homozygous lof variants)\npRec - the probability of being intolerant of homozygous, but not heterozygous lof variants\npNull - the probability of being tolerant of both heterozygous and homozygous lof variants";

my @CommentGnomadGenome = ('gnomAD_genome_ALL',
                           'gnomAD_genome_AFR',
                           'gnomAD_genome_AMR',
                           'gnomAD_genome_ASJ',
                           'gnomAD_genome_EAS',
                           'gnomAD_genome_FIN',
                           'gnomAD_genome_NFE',
                           'gnomAD_genome_OTH');





my @CommentGnomadExome = ('gnomAD_exome_ALL',
                          'gnomAD_exome_AFR',
                          'gnomAD_exome_AMR',
                          'gnomAD_exome_ASJ',
                          'gnomAD_exome_EAS',
                          'gnomAD_exome_FIN',
                          'gnomAD_exome_NFE',
                          'gnomAD_exome_OTH');


my %CommentInterVar = (
'PVS1' => "Certain types of variants (e.g., nonsense, frameshift, canonical +- 1 or 2 splice sites, initiation codon, single exon or multiexon deletion) in a gene where LOF is a known mechanism of diseas",
'PS1' => "Same amino acid change as a previously established pathogenic variant regardless of nucleotide change
    Example: Val->Leu caused by either G>C or G>T in the same codon",
'PS2' => "De novo (both maternity and paternity confirmed) in a patient with the disease and no family history",
'PS3' => "Well-established in vitro or in vivo functional studies supportive of a damaging effect on the gene or gene product",
'PS4' => "The prevalence of the variant in affected individuals is significantly increased compared with the prevalence in controls; OR>5 in all the gwas, the dataset is from gwasdb jjwanglab.org/gwasdb",
'PM1' => "Located in a mutational hot spot and/or critical and well-established functional domain (e.g., active site of an enzyme) without benign variation",
'PM2' => "Absent from controls (or at extremely low frequency if recessive) (Table 6) in Exome Sequencing Project, 1000 Genomes Project, or Exome Aggregation Consortium",
'PM3' => "For recessive disorders, detected in trans with a pathogenic variant",
'PM4' => "Protein length changes as a result of in-frame deletions/insertions in a nonrepeat region or stop-loss variants",
'PM5' => "Novel missense change at an amino acid residue where a different missense change determined to be pathogenic has been seen before;Example: Arg156His is pathogenic; now you observe Arg156Cys",
'PM6' => "Assumed de novo, but without confirmation of paternity and maternity",
'PP1' => "Cosegregation with disease in multiple affected family members in a gene definitively known to cause the disease",
'PP2' => "Missense variant in a gene that has a low rate of benign missense variation and in which missense variants are a common mechanism of disease",
'PP3' => "Multiple lines of computational evidence support a deleterious effect on the gene or gene product (conservation, evolutionary, splicing impact, etc.) sfit for conservation, GERP++_RS for evolutionary, splicing impact from dbNSFP",
'PP4' => "Patient's phenotype or family history is highly specific for a disease with a single genetic etiology",
'PP5' => "Reputable source recently reports variant as pathogenic, but the evidence is not available to the laboratory to perform an independent evaluation",
'BA1' => "BA1 Allele frequency is >5% in Exome Sequencing Project, 1000 Genomes Project, or Exome Aggregation Consortium",
'BS1' => "Allele frequency is greater than expected for disorder (see Table 6) > 1% in ESP6500all ExAc? need to check more",
'BS2' => "Observed in a healthy adult individual for a recessive (homozygous), dominant (heterozygous), or X-linked (hemizygous) disorder, with full penetrance expected at an early age, check ExAC_ALL",
'BS3' => "Well-established in vitro or in vivo functional studies show no damaging effect on protein function or splicing",
'BS4' => "Lack of segregation in affected members of a family",
'BP1' => "Missense variant in a gene for which primarily truncating variants are known to cause disease truncating:  stop_gain / frameshift deletion/  nonframshift deletion
	    We defined Protein truncating variants  (4) (table S1) as single-nucleotide variants (SNVs) predicted to introduce a premature stop codon or to disrupt a splice site, small insertions or deletions (indels) predicted to disrupt a transcript reading frame, and larger deletions ",
'BP2' => "Observed in trans with a pathogenic variant for a fully penetrant dominant gene/disorder or observed in cis with a pathogenic variant in any inheritance pattern",
'BP3' => "In-frame deletions/insertions in a repetitive region without a known function if the repetitive region is in the domain, this BP3 should not be applied.",
'BP4' => "Multiple lines of computational evidence suggest no impact on gene or gene product (conservation, evolutionary,splicing impact, etc.)",
'BP5' => "Variant found in a case with an alternate molecular basis for disease.
check the genes whether are for mutilfactor disorder. The reviewers suggeset to disable the OMIM morbidmap for BP5",
'BP6' => "Reputable source recently reports variant as benign, but the evidence is not available to the laboratory to perform an independent evaluation; Check the ClinVar column to see whether this is \"benign\".",
'BP7' => "A synonymous (silent) variant for which splicing prediction algorithms predict no impact to the 
    splice consensus sequence nor the creation of a new splice site AND the nucleotide is not highly conserved"
);


my @CommentFunc = ( 	'ExonicFunc.refGene',
			'AAChange.refGene',
			'GeneDetail.refGene');


my @CommentPhenotype = ( 'Disease_description.refGene');


#########FILLING COLUMN TITLES FOR SHEETS
$worksheet->write_row( 0, 0, \@columnTitles );
$worksheetACMG->write_row( 0, 0, \@columnTitles );
		
#write comment for pLI
$worksheet->write_comment( 0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3 );
$worksheetACMG->write_comment( 0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3  );


#FILLING COLUMN TITLES FOR TRIO SHEETS
if (defined $trio){
	$worksheetHTZcompo->write_row( 0, 0, \@columnTitles );
	$worksheetAR->write_row( 0, 0, \@columnTitles );
	$worksheetSNPmumVsCNVdad->write_row( 0, 0, \@columnTitles );
	$worksheetSNPdadVsCNVmum->write_row( 0, 0, \@columnTitles );
	$worksheetDENOVO->write_row( 0, 0, \@columnTitles );

			
	#write pLI comment
	$worksheetHTZcompo->write_comment(0,$dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);
	$worksheetAR->write_comment(0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);
	$worksheetSNPmumVsCNVdad->write_comment(0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);
	$worksheetSNPdadVsCNVmum->write_comment(0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);
	$worksheetDENOVO->write_comment( 0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);

}

#FILLING COLUMN TITLES FOR CANDIDATES SHEET

if($candidates ne ""){
	$worksheetCandidats->write_row( 0, 0, \@columnTitles );
	$worksheetCandidats->write_comment(0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3 );
}


####################################################################
#
#
#
#

my %dicoGeneForHTZcompo;
my $previousGene ="";
my $filterBool;
$dicoGeneForHTZcompo{$previousGene}{'ok'}=0;
$dicoGeneForHTZcompo{$previousGene}{'cnt'}=0;

#############################################
##################   Start parsing VCF

open( VCF , "<$incfile" )or die("Cannot open vcf file ".$incfile) ;


while( <VCF> ){
  	$current_line = $_;

#############################################
##############   skip header
	next if ($current_line=~/^##/);

	chomp $current_line;
	@line = split( /\t/, $current_line );	
	
	#DEBUG print STDERR $dicoColumnNbr{'Gene.refGene'}."\n";



#############################################
##############   Treatment for First line to create header of the output

	if ( $line[0] eq "#CHROM" )   {
	


		#############CHECK IF $nbrSample == $cmpt else exit error nbr sample
		
	    #find correct index for samples of the family		
		#for each sample position   	
		for( my $sampleIndex = 9 ; $sampleIndex < scalar @line; $sampleIndex++){
			
			#for each final position for sample 
			foreach my $finCol (keys %dicoSamples){
    	    #DEBUG
					
				if ($dicoSamples{$finCol}{'columnName'} eq "Genotype-".$line[$sampleIndex] ){
					$dicoSamples{$finCol}{'columnIndex'} =  $sampleIndex;				
					last;
				}
			}

#			foreach my $name (@sampleList)
				#if($line[$sampleIndex]
				
			
		}	
	
    

######### Increase row number 
	    	$worksheetLine ++;
	    	$worksheetLineACMG ++;

		if (defined $trio){

			$worksheetLineHTZcompo ++;
			$worksheetLineHTZcompo ++;
			$worksheetLineAR ++; 
			$worksheetLineSNPmumVsCNVdad ++;
			$worksheetLineSNPdadVsCNVmum ++;
			$worksheetLineDENOVO ++;

		}

		if($candidates ne ""){
			$worksheetLineCandidats ++;
		}


		next;
		
#############################################
##############################
##########  start to compute variant lines	

	}else {



		#initialise final printable string
		@finalSortData = ("");
		
		my $alt="";
		my $ref="";

		#Split line with tab 
					
		#DEBUG		print $current_line,"\n";	
		
		#filling output line with classical first vcf columns
		$finalSortData[$dicoColumnNbr{'#CHROM'}]=	$line[0];
		$finalSortData[$dicoColumnNbr{'POS'}]=		$line[1];
		$finalSortData[$dicoColumnNbr{'ID'}]=		$line[2];
		$finalSortData[$dicoColumnNbr{'REF'}]=		$line[3];
		$finalSortData[$dicoColumnNbr{'ALT'}]=		$line[4];
		$finalSortData[$dicoColumnNbr{'FILTER'}]=	$line[6];




#############################################
########### Split INFOS #####################
	
		my %dicoInfo;
		my @infoList = split(';', $line[7] );	
		foreach my $info (@infoList){
			my @infoKeyValue = split('=', $info );
			if (scalar @infoKeyValue == 2){
				
				$infoKeyValue[1] =~ s/\\x3d/=/g;
				$infoKeyValue[1] =~ s/\\x3b/;/g;
				
				$dicoInfo{$infoKeyValue[0]} = $infoKeyValue[1];
				#DEBUG
				#print $infoKeyValue[1]."\n";
			}
		}

#DEBUG
#print Dumper(\%dicoInfo);


		#select only x% pop freq 
		#Use pop freq threshold as an input parameter (default = 1%)
		next if(( $dicoInfo{'gnomAD_genome_ALL'} ne ".") && ($dicoInfo{'gnomAD_genome_ALL'} >= $popFreqThr));  
	
		
		#FILTERING according to newHope option and filterList option
		$filterBool = 0;
		if(defined $newHope){
			#Keep only NON PASS or PASS + MPA ranking = 8
			#next if ($finalSortData[$dicoColumnNbr{'FILTER'}] eq "PASS" && $dicoInfo{'MPA_ranking'} < 8);
			switch ($finalSortData[$dicoColumnNbr{'FILTER'}]){
				case (\@filterArray) { $filterBool=1 };
			}
			next if($filterBool == 1);
			next if($dicoInfo{'MPA_ranking'} < 8 );

		}else{
			#remove NON PASS variant and remove MPA_Ranking = 8

			switch ($finalSortData[$dicoColumnNbr{'FILTER'}]){
				case (\@filterArray) {$filterBool=0}
				else {$filterBool=1}
			}
			next if ($filterBool == 1);
			next if( $dicoInfo{'MPA_ranking'} == 8);
			#next if ($finalSortData[$dicoColumnNbr{'FILTER'}] ne "PASS");
		
		}



		#filling output line, check if INFO exists in the VCF
		foreach my $keys (sort keys %dicoColumnNbr){
			
#			print "keysListe\t#".$keys."#\n";

			if (defined $dicoInfo{$keys}){
				
				$finalSortData[$dicoColumnNbr{$keys}] = $dicoInfo{$keys};
				#DEBUG
				#print "finalSort\t".$finalSortData[$dicoColumnNbr{$keys}]."\n";
				#print "dicoInfo\t".$dicoInfo{$keys}."\n";
				#print "keys\t".$keys."\n";
			}
		}	


		#Phenolyzer Column
		if($phenolyzerFile ne ""){

			my @geneList = split(';', $finalSortData[$dicoColumnNbr{'Gene.refGene'}] );	
			foreach my $geneName (@geneList){
				#if (defined $phenolyzerGene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]} )
				if (defined $phenolyzerGene{$geneName} ){
					$finalSortData[$dicoColumnNbr{'Phenolyzer'}] .= $phenolyzerGene{$geneName}{'Raw'}."\t";
				}

			}	

		}


		#CNV GENES
		if($cnvGeneList ne ""){
	
			my @geneList = split(';', $finalSortData[$dicoColumnNbr{'Gene.refGene'}] );	
			foreach my $geneName (@geneList){
				#print $geneName."\n";
				if (defined $cnvGene{$geneName} ){
					#print $geneName."toto\n";
					$finalSortData[$dicoColumnNbr{'SecondHit-CNV'}] .= $cnvGene{$geneName}."\n";
					#print $finalSortData[$dicoColumnNbr{'SecondHit-CNV'}]."\n";
				}

			}	
		}





		#MPA COMMENT
		#create string with array
		$commentMPAscore = "";
		foreach my $keys (@CommentMPA_score){
			if (defined $dicoInfo{$keys} ){
				$commentMPAscore .= $keys."\t= ".$dicoInfo{$keys}."\n";
			}else{
				$commentMPAscore .= $keys."\n";
			}
			
			#refine MPA_rank for rank 7 missense with MPA final score  or  for other ranking with pLI
			if($keys eq "MPA_final_score" &&  $finalSortData[$dicoColumnNbr{'MPA_ranking'}] == 7 ){
				$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += (10-$dicoInfo{$keys})/100;
				#print $finalSortData[$dicoColumnNbr{'MPA_ranking'}]."\n";
				
			}

		}

		#refine MPA_rank  with pLI
		if(defined $dicoInfo{"pLi.refGene"} && $dicoInfo{"pLi.refGene"} ne "." ){
				$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += 0.001-(($dicoInfo{"pLi.refGene"}/1000 +  $dicoInfo{"pRec.refGene"}/10000 + $dicoInfo{"pNull.refGene"}/100000)) ;
				#print $finalSortData[$dicoColumnNbr{'MPA_ranking'}]."\n";
			
		}else{
			$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += 0.001;
		}
		
		#GNOMAD_EXOME COMMENT
		#create string with array
		$commentGnomADExomeScore = "";
		foreach my $keys (@CommentGnomadExome){
			if (defined $dicoInfo{$keys} ){
				$commentGnomADExomeScore .= $keys."\t= ".$dicoInfo{$keys}."\n";
			}
		}
		
		#GNOMAD_GENOME COMMENT
		#create string with array
		$commentGnomADGenomeScore = "";
		foreach my $keys (@CommentGnomadGenome){
			if (defined $dicoInfo{$keys} ){
				$commentGnomADGenomeScore .= $keys."\t= ".$dicoInfo{$keys}."\n";
			}
		}

		#INTERVAR COMMENT
		#create string with hash
		$commentInterVar = "";
		foreach my $keys (keys %CommentInterVar){
			if (defined $dicoInfo{$keys} && $dicoInfo{$keys} ne "." && $dicoInfo{$keys} ne "0"){
				$commentInterVar .= $keys."\t= ".$dicoInfo{$keys}."\n".$CommentInterVar{$keys}."\n\n";
			}
		}


		#FUNCTION REFGENE COMMENT
		#create string with array
		$commentFunc = "";
		foreach my $keys (@CommentFunc){
			if (defined $dicoInfo{$keys} ){
				$dicoInfo{$keys} =~ s/;/\n /g;
				$commentFunc .= $keys.":\n ".$dicoInfo{$keys}."\n\n";
			}
		}

		#PHENOTYPES REFGENE COMMENT (OMIM)
		#create string with array
		$commentPhenotype = "";
		foreach my $keys (@CommentPhenotype){
			if (defined $dicoInfo{$keys} ){

				$dicoInfo{$keys} =~ s/_DISEASE:/\n\nDISEASE:/g;
				$commentPhenotype .= $keys.":\n ".$dicoInfo{$keys}."\n\n";
			}
		}





############################################
#############	Check FORMAT related to caller
		#
		#		GT:AD:DP:GQ:PL (haplotype caller)
		#		GT:DP:GQ  => multiallelic line , vcf not splitted => should be done before, STOP RUN?
		#		GT:GOF:GQ:NR:NV:PL (platyplus caller)
		if($line[8] eq "GT:AD:DP:GQ:PL" || $line[8] eq "GT:AB:AD:DP:GQ:PL"){
			$caller = "GATK";
		}elsif($line[8] eq "GT:GOF:GQ:NR:NV:PL"){
			$caller = "platypus";
		}elsif($line[8] eq "GT:DP:GQ"){
			$caller = "other _ GT:DP:GQ";
			#print STDERR "Multi-allelic line detection. Please split the vcf file in order to get 1 allele by line\n";
			#print STDERR $current_line ."\n";
			#exit 1; 
		}else{
			print STDERR "The Format of the Caller used for this line is unknown. Processing is a risky business. This line won't be processed.\n";
			print STDERR $current_line ."\n";
			next;
			#exit 1; 
		}
		

############################################
#############   Parse Genotypes
		#
		#genotype concatenation for easy hereditary status
		$familyGenotype = "_";		
		$commentGenotype = "CALLER = ".$caller."\t QUALITY = ".$line[5]."\n\n";


		#for each sample sort by sample wanted 
		foreach my $finalcol ( sort {$a <=> $b}  (keys %dicoSamples) ) {

			#DEBUG print "tata\t".$finalcol."\n";
			#DEBUG	print $line[$dicoSamples{$finalcol}{'columnIndex'}]."\n";
				
				my @genotype = split(':', $line[$dicoSamples{$finalcol}{'columnIndex'}] );

				my $DP;		#total Depth
				my $adalt;	#Alternative Allelic Depth
				my $adref;	#Reference Allelic Depth
				my $AB;		#Allelic balancy
				my $AD;		#Final Allelic Depth


				if (scalar @genotype > 1 && $caller ne ""){

					if(	$caller eq "GATK"){

						#check if variant is not called
						if ($genotype[2] eq "."){
							$DP = 0;
							$AD = "0,0";

						}elsif (length $line[8] == 14){
							
							$DP = $genotype[2];
							$AD = $genotype[1];
						
						}else{
							$DP = $genotype[3];
							$AD = $genotype[2];
						}

					}elsif($caller eq "platypus"){
						
						if($genotype[3] =~ m/,/){
							my @genotype_DPsplit = split(',', $genotype[3]);
							my @genotype_ADsplit = split(',', $genotype[4]);
							$DP = $genotype_DPsplit[0];
							$AD = ($genotype_DPsplit[0] - $genotype_ADsplit[0]);
							$AD .= ",".$genotype_ADsplit[0];
							
						}else{
							$DP = $genotype[3];
							$AD = ($genotype[3] - $genotype[4]);
							$AD .= ",".$genotype[4];
						}
					}elsif($caller eq "other _ GT:DP:GQ"){
						
						$DP = $genotype[1];
						$AD = 0;
						$AD .= ",0";
					}

					
					#DEBUG
					#print $genotype[2]."\n";


					#split allelic depth
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


				#DEBUG print STDERR "indexSample\t".$dicoSamples{$finalcol}{'columnIndex'}."\n";

				#put the genotype and comments info into string
				$finalSortData[$dicoColumnNbr{$dicoSamples{$finalcol}{'columnName'}}] = $genotype[0];
				$commentGenotype .=  $dicoSamples{$finalcol}{'columnName'}."\t -\t ".$genotype[0]."\nDP = ".$DP."\t AD = ".$AD."\t AB = ".$AB."\n\n";
				$familyGenotype .= $genotype[0]."_";
		
		} #END of Sample Treatment	


		
		#	print Dumper(\@finalSortData);

#############################################################################
#########FILL HASH STRUCTURE FOR FINAL SORT AND OUTPUT, according to rank
		
		#concatenate chrom_POS_REF_ALT to get variant ID
		$variantID = $line[0]."_".$line[1]."_".$line[3]."_".$line[4]."_".$caller;
		
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'finalArray'} = [@finalSortData] ; 
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentGnomADexome'} = $commentGnomADExomeScore  ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentGnomADgenome'} = $commentGnomADGenomeScore ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentGenotype'} = $commentGenotype ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentMPAscore'} = $commentMPAscore  ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentInterVar'} = $commentInterVar  ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentFunc'} = $commentFunc  ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentPhenotype'} = $commentPhenotype  ;

			#initialize worksheet
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} = "";

			#test multiple geneNames
			my @geneList = split(';', $finalSortData[$dicoColumnNbr{'Gene.refGene'}] );	
			foreach my $geneName (@geneList){

				#ACMG
				#if(defined $ACMGgene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]} )
				if(defined $ACMGgene{$geneName} ){
					$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#ACMG";
				}
		
				#CANDIDATES
				if($candidates ne ""){
					#if(defined $candidateGene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]} )
					if(defined $candidateGene{$geneName} ){
						$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#CANDIDATES";
				
					}
				
				}

				#PHENOLYZER COMMENT
				$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentPhenolyzer'} = "";
				
				#if(defined  $phenolyzerGene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]})
				if(defined  $phenolyzerGene{$geneName}){
					$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentPhenolyzer'} .= $phenolyzerGene{$geneName}{'comment'}."\n\n"  ;

				}

		
			}







##########additionnal analysis in TRIO context according to family genotype
			if (defined $trio){
			
				switch ($familyGenotype){
					#Find de novo
					case ["_1/1_0/0_0/0_" ,"_0/1_0/0_0/0_","_1/0_0/0_0/0_"] {$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#DENOVO";}

					#find Autosomique Recessive
					case ["_1/1_0/1_0/1_" , "_1/1_1/0_1/0_"] {$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#AUTOREC";}

				#Find SNVvsCNV

					case ["_1/1_0/0_0/1_" , "_1/1_0/0_1/0_"] {$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#SNPmCNVp";}
					case ["_1/1_0/1_0/0_" , "_1/1_1/0_0/0_"] {$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#SNPpCNVm";}

				}


			}




##########create pLI comment and format
		if(defined $dicoInfo{'pLi.refGene'} && $dicoInfo{'pLi.refGene'} ne "." ){

			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} =  "pLI = ".$dicoInfo{'pLi.refGene'}."\npRec = ".$dicoInfo{'pRec.refGene'}."\npNull = ".$dicoInfo{'pNull.refGene'} ."\n\n";
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  sprintf('#%2.2X%2.2X%2.2X',($dicoInfo{'pLi.refGene'}*255 + $dicoInfo{'pRec.refGene'}*255),($dicoInfo{'pRec.refGene'}*255 + $dicoInfo{'pNull.refGene'} * 255),0) ;
        		
			
        		$format_pLI = $workbook->add_format(bg_color => $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'});

			if(defined $dicoInfo{'Missense_Z_score.refGene'} && $dicoInfo{'Missense_Z_score.refGene'} ne "." ){
					$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} .=  "Missense Z-score = ".$dicoInfo{'Missense_Z_score.refGene'}."\n\n";
			}
		
		}else{	
			
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} =  "." ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  '#FFFFFF' ;
			
			#$format_pLI = $workbook->add_format(bg_color => '#FFFFFF');
		}


##########Add function and expression infos in comment
		if(defined $dicoInfo{'Function_description.refGene'}  && $dicoInfo{'Function_description.refGene'} ne "." ){
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} .= "Function Description:\n".$dicoInfo{'Function_description.refGene'}."\n\n"; 
			
			if(defined $dicoInfo{'Tissue_specificity(Uniprot).refGene'} && $dicoInfo{'Tissue_specificity(Uniprot).refGene'} ne "."  ){
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} .= "Tissue specificity:\n".$dicoInfo{'Tissue_specificity(Uniprot).refGene'}."\n"; 	
			} 

		} 
		

	} #END of IF-ELSE(#CHROM)	





#DEBUG			print Dumper(\%dicoColumnNbr);
##############check hereditary hypothesis or genes to fill sheets





############ TIME TO FILL THE XLSX OUTPUT FILE




###############	#additionnal analysis in TRIO context
			if (defined $trio){

						
				#DEBUG	print $familyGenotype."\t".$format_pLI."\n";

				#Find HTZ composite
				if ($familyGenotype eq "_0/1_0/1_0/0_" || $familyGenotype eq "_0/1_0/0_0/1_"){	
					if(defined $dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]} ){
						
						$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'cnt'} ++;

						if($familyGenotype =~ /0\/0_$/){

							$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'pvsm'} ++;

							if( $dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'mvsp'} >= 1){  
								$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'ok'} = 1;
							}
						}else{
							$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'mvsp'} ++;

							if( $dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'pvsm'} >= 1){  
								$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'ok'} = 1;
							}
						}
						
						#erase line with empty array
						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@emptyArray );
#						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@finalSortData );

						#create reference of Hashes
						my $hashTemp = $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID};
						my $hashColumn_ref = \%dicoColumnNbr;

						#write line in the HTZcompo sheet
						writeThisSheet ($worksheetHTZcompo,
										$worksheetLineHTZcompo,
										$format_pLI,
										$case,
										$hashTemp,
										$hashColumn_ref
									);
						
						$worksheetLineHTZcompo ++;

					}else{

#						if(defined $dicoGeneForHTZcompo{$previousGene} && $dicoGeneForHTZcompo{$previousGene}{'ok'}==0 ){
						if(($previousGene ne $finalSortData[$dicoColumnNbr{'Gene.refGene'}]) && $dicoGeneForHTZcompo{$previousGene}{'ok'}==0 ){
							$worksheetLineHTZcompo -= $dicoGeneForHTZcompo{$previousGene}{'cnt'};
						}
						$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'cnt'} = 1;
						$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'ok'} = 0;

						if($familyGenotype =~ /0\/0_$/){

							$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'pvsm'} = 1;
							$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'mvsp'} = 0;

						}else{

							$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'mvsp'} = 1;
							$dicoGeneForHTZcompo{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'pvsm'} = 0;

						}

						#erase line with empty array
						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@emptyArray );
						#$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@finalSortData );
						#create reference of Hashes
						my $hashTemp = $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID};
						my $hashColumn_ref = \%dicoColumnNbr;

						#write line in the HTZcompo sheet
						writeThisSheet ($worksheetHTZcompo,
										$worksheetLineHTZcompo,
										$format_pLI,
										$case,
										$hashTemp,
										$hashColumn_ref
									);
																																									$worksheetLineHTZcompo ++;
																																									$previousGene = $finalSortData[$dicoColumnNbr{'Gene.refGene'}];
						
						}

					}#END IF HTZ COMPO genotype


#				}			
			}# END IF TRIO

}#END WHILE VCF

#########################################################################
#################### Sort by MPA ranking for the output


foreach my $rank (sort {$a <=> $b} keys %hashFinalSortData){
	#print $rank."\n";
	
	foreach my $variant ( keys %{$hashFinalSortData{$rank}}){

		#print $variant."\n";
#		$format_pLI = $workbook->add_format(bg_color => '#FFFFFF');

		#create reference of Hashes
		my $hashTemp = $hashFinalSortData{$rank}{$variant};
		my $hashColumn_ref = \%dicoColumnNbr;

##############################################################
###########################      ALL     #####################

		writeThisSheet ($worksheet,
						$worksheetLine,
						$format_pLI,
						$case,
						$hashTemp,
						$hashColumn_ref
					);
		$worksheetLine ++;
	
#						$hashFinalSortData{$rank}{$variant}{'commentMPAscore'},
#						$hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'},
#						$hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'},
#						$hashFinalSortData{$rank}{$variant}{'commentGenotype'},
#						$hashFinalSortData{$rank}{$variant}{'commentFunc'},
#						$hashFinalSortData{$rank}{$variant}{'commentPhenotype'},
#						$hashFinalSortData{$rank}{$variant}{'commentPhenolyzer'},
#						$hashFinalSortData{$rank}{$variant}{'commentInterVar'},
#						$hashFinalSortData{$rank}{$variant}{'commentpLI'},
#						$hashFinalSortData{$rank}{$variant}{'colorpLI'},
#						$hashFinalSortData{$rank}{$variant}{'finalArray'},
#						%dicoColumnNbr





##############################################################
################# ACMG DS #############

		if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /ACMG/){

			writeThisSheet ($worksheetACMG,
							$worksheetLineACMG,
							$format_pLI,
							$case,
							$hashTemp,
							$hashColumn_ref
						);
			
			$worksheetLineACMG ++;
		}
	


##############################################################
################ CANDIDATES #############

		if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /CANDIDATE/){

			writeThisSheet ($worksheetCandidats,
							$worksheetLineCandidats,
							$format_pLI,
							$case,
							$hashTemp,
							$hashColumn_ref
						);
			
			
			$worksheetLineCandidats ++;
		}
	
		


##############################################################
#################### TRIO ################
##############################################################
		
		if(defined $trio){
			
##############################################################
################ DENOVO  ######################
			if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /DENOVO/){

				writeThisSheet ($worksheetDENOVO,
								$worksheetLineDENOVO,
								$format_pLI,
								$case,
								$hashTemp,
								$hashColumn_ref
							);
				$worksheetLineDENOVO ++;

				next;

			}	


##############################################################
################  AUTOSOMIC RECESSIVE HOMOZYGOUS ###############
			if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /AUTOREC/){

				writeThisSheet ($worksheetAR,
								$worksheetLineAR,
								$format_pLI,
								$case,
								$hashTemp,
								$hashColumn_ref
							);
				$worksheetLineAR ++;

				next;

			}

##############################################################
################ SNPvsCNV ##########################
			if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /SNPpCNVm/) {

				writeThisSheet ($worksheetSNPdadVsCNVmum,
								$worksheetLineSNPdadVsCNVmum,
								$format_pLI,
								$case,
								$hashTemp,
								$hashColumn_ref
							);
				
				$worksheetLineSNPdadVsCNVmum ++;
				next;

			}	


			if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /SNPmCNVp/){

				writeThisSheet ($worksheetSNPmumVsCNVdad,
								$worksheetLineSNPmumVsCNVdad,
								$format_pLI,
								$case,
								$hashTemp,
								$hashColumn_ref
							);
				$worksheetLineSNPmumVsCNVdad ++;
				
				next;

			}	


		}   # END IF TRIO
	
	}   # END FOREACH VARIANT
}	#END FOREACH RANK



#############add autofilters within the sheets

$worksheet->autofilter('A1:Z'.$worksheetLine); # Add autofilter until the end

if(defined $trio){

	$worksheetHTZcompo->autofilter('A1:Z'.$worksheetLineHTZcompo); # Add autofilter
  
	$worksheetAR->autofilter('A1:Z'.$worksheetLineAR); # Add autofilter
  
	$worksheetSNPmumVsCNVdad->autofilter('A1:Z'.$worksheetLineSNPmumVsCNVdad); # Add autofilter
  
	$worksheetSNPdadVsCNVmum->autofilter('A1:Z'.$worksheetLineSNPdadVsCNVmum); # Add autofilter
  
	$worksheetDENOVO->autofilter('A1:Z'.$worksheetLineDENOVO); # Add autofilter


}

if($candidates ne ""){
			$worksheetCandidats->autofilter('A1:Z'.$worksheetLineCandidats);
}

$workbook->close();
close(VCF);

print STDERR "Done!\n\n\n";

###################################################################
######################### FUNCTIONS ###############################
# Write in sheets
sub writeThisSheet {
	my ($worksheet, 
		$worksheetLine, 
		$format_pLI,	
		$case,
		$hashTemp_ref, 
		$hashColumn_ref	) = @_;

	my %hashTemp = %{$hashTemp_ref};  #Dereference the hash
	my %hashColumn = %{$hashColumn_ref};

#						$hashTemp{'commentMPAscore'},
#						$hashTemp{'commentGnomADgenome'},
#						$hashTemp{'commentGnomADgenome'},
#						$hashTemp{'commentGenotype'},
#						$hashTemp{'commentFunc'},
#						$hashTemp{'commentPhenotype'},
#						$hashTemp{'commentPhenolyzer'},
#						$hashTemp{'commentInterVar'},
#						$hashTemp{'commentpLI'},
#						$hashTemp{'colorpLI'},
#						$hashTemp{'finalArray'},





	#dEBUG
#	print  "casecasecasecase         ". $hashTemp{'commentMPAscore'}."____".$hashColumn{"Gene.refGene"}."\n\n"; 


			$worksheet->write_row( $worksheetLine, 0, $hashTemp{'finalArray'} );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'MPA_ranking'},		$hashTemp{'commentMPAscore'} ,x_scale => 2, y_scale => 5 );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'gnomAD_genome_ALL'},	$hashTemp{'commentGnomADgenome'} ,x_scale => 3, y_scale => 2  );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'gnomAD_exome_ALL'},	$hashTemp{'commentGnomADexome'} ,x_scale => 3, y_scale => 2  );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'Genotype-'.$case},	$hashTemp{'commentGenotype'} ,x_scale => 2, y_scale => 3 );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'Func.refGene'},		$hashTemp{'commentFunc'} ,x_scale => 3, y_scale => 3  );
			
			if ($hashTemp{'commentPhenotype'} ne ""){

				$worksheet->write_comment( $worksheetLine,$hashColumn{'Phenotypes.refGene'}, $hashTemp{'commentPhenotype'} ,x_scale => 7, y_scale => 5  );
			}
			
			if ($hashTemp{'commentPhenolyzer'} ne ""){
				$worksheet->write_comment( $worksheetLine,$hashColumn{'Phenolyzer'}, $hashTemp{'commentPhenolyzer'} ,x_scale => 2 );
			}

			if ($hashTemp{'commentInterVar'} ne ""){
				$worksheet->write_comment( $worksheetLine,$hashColumn{'InterVar_automated'}, $hashTemp{'commentInterVar'} ,x_scale => 7, y_scale => 5  );
			}


			if ($hashTemp{'commentpLI'} ne "."){

        		$format_pLI = $workbook->add_format(bg_color => $hashTemp{'colorpLI'});


				$worksheet->write( $worksheetLine,$hashColumn{'Gene.refGene'}, $hashTemp{'finalArray'}[$hashColumn{'Gene.refGene'}]     ,$format_pLI );
				$worksheet->write_comment( $worksheetLine,$hashColumn{'Gene.refGene'},$hashTemp{'commentpLI'},x_scale => 5, y_scale => 5  );
			}	

}#END OF SUB

exit 0;
