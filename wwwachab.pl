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
use Data::Dumper;

#parameters
my $man = "USAGE : \nperl achab.pl 
\n--vcf <vcf_file> 
\n--outDir <output directory (default = current dir)> 
\n--outPrefix <output file prelifx (default = \"\")> 
\n--candidates <file with gene symbol of interest>  
\n--phenolyzerFile <phenolyzer output file suffixed by predicted_gene_scores>   
\n--popFreqThr <allelic frequency threshold from 0 to 1 default=0.01> 
\n--trio (requires case dad and mum option to be filled) 
\n\t--case <index_sample_name> 
\n\t--dad <father_sample_name> 
\n\t--mum <mother_sample_name>  
\n--customInfoList  <comma separated list of vcf-info names (will be added in a new column)>  
\n--filterList <comma separated list of VCF FILTER to output (default='PASS', included )>   
\n--cnvGeneList <File with gene symbol + annotation (1 tab separated), involved by parallel CNV calling >
\n--customVCF <VCF format File with custom annotation (if variant matches then INFO field annotations will be added in new column)>
\n--mozaicRate <mozaic rate value from 0 to 1, it will color 0/1 genotype according to this value  (default=0.2 as 20%)>
\n--mozaicDP <ALT variant Depth, number of read supporting ALT, it will give darker color to the 0/1 genotype  (default=5)>
\n--newHope (only popFreqThr filter is applied (no more FILTER or MPA_ranking))>
\n--affected <comma separated list of samples affected by phenotype (assuming they support the same genotype >
\n--favouriteGeneRef <File with transcript references to extract in a new column (1 transcript by line) >
\n--filterCustomVCF <integer value, penalizes variant if its frequency in the customVCF is greater than [value] (default key of info field : found=[value])  >
\n--filterCustomVCFRegex <string pattern used as regex to search for a specific field to filter customVCF (default key of info field : 'found=')  >";


#################################### VARIABLES INIT ########################

#catch argument for METADATA 
my $achabArg = ""; 
foreach my $a(@ARGV){
    $achabArg .= $a."   ";
}



my $help;
my $current_line;
my $incfile = "";
my $outDir = ".";
my $outPrefix = "";
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
my $mozaicRate = "" ;
my $mozaicDP = "";


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

my $customVCF_File = "";
my $customVCF_Line;
my @customVCF_List;
my %customVCF_variant;
#threshold to filter out bias related frequent variants 
my $filterCustomVCF = "";
my $filterCustomVCFRegex = "";

#vcf parsing and output
my @line;
my $variantID; 
my @geneListTemp;
my @geneList;
my $mozaicSamples = "";
my $count = 0;

#Data structure
my @finalSortData;
my $familyGenotype;		
my %hashFinalSortData;

my $worksheetTAG = "";

# favourite NM refGene
my $favouriteGeneRef = "";
my @geneRefArray;
my %geneRef_gene;
my $geneRef_line;

#affected samples
my $affected = "";    # next if affected_sample = case or dad or mum
my @affectedArray;
my %hashAffected;
my @nonAffectedArray;

#Variable for genotype checking 
my @strangerNULL;
my @strangerREF;
my @strangerHTZ;
my @strangerHMZ;


# METADATA

# inheritance checking test
my $dadVariant = 0;
my $mumVariant = 0;
my $caseDadVariant = 0;
my $caseMumVariant = 0;

my $vcfHeader = "";

my @buttonArray ;
my %tagsHash;

#$arguments = GetOptions( "vcf=s" => \$incfile ) or pod2usage(-vcf => "$0: argument required\n") ;

GetOptions( 	"vcf=s"				=> \$incfile,
		"case=s"			=> \$case,
		"dad=s"				=> \$dad, 
		"mum=s"				=> \$mum, 
		"trio"				=> \$trio,
		"candidates:s"			=> \$candidates,
		"outDir=s"			=> \$outDir,
		"outPrefix:s"			=> \$outPrefix,
		"phenolyzerFile:s"		=> \$phenolyzerFile,
		"popFreqThr=s"			=> \$popFreqThr, 
		"customInfoList:s"		=> \$customInfoList, 
		"filterList:s"			=> \$filterList,				
		"cnvGeneList:s"			=> \$cnvGeneList,
		"customVCF:s"			=> \$customVCF_File,
		"mozaicRate:s"			=> \$mozaicRate,
		"mozaicDP:s"			=> \$mozaicDP,
		"newHope"			=> \$newHope,
		"favouriteGeneRef:s"			=> \$favouriteGeneRef,
		"affected:s"		=> \$affected,
		"filterCustomVCF:s"			=> \$filterCustomVCF,
		"filterCustomVCFRegex:s"	=>	\$filterCustomVCFRegex,
		"help|h"			=> \$help);
				
				


#check mandatory arguments
if(defined $help || $incfile eq ""){
	die("$man");
}

#add underscore to output prefix
if($outPrefix ne ""){
	$outPrefix .= "_";
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


#check mozaic parameters
if($mozaicRate eq ""){
	$mozaicRate = 0.2;
}
if($mozaicDP eq ""){
	$mozaicDP = 5;
}


#check sample list param
if(defined $trio && ($case eq "" || $dad eq "" || $mum eq "")){	
	die("TRIO option requires 3 sample names. Please, give --case, --dad and --mum sample name arguments.\n");
}
			
#TODO affected samples
#define affected samples List
if($affected ne ""){
	chomp $affected;
	@affectedArray = split(/,/ , $affected);
	foreach my $affSample (@affectedArray){
		if (defined $hashAffected{$affSample}){
				#nothing to do
		}else{	
			$hashAffected{$affSample} = "affected";
		}
	}
}




print  STDERR "Starting a new fishing trip ... \n" ; 
print  STDERR "Hope we will catch-a-lot ... \n" ; 
print  STDERR "Processing vcf file ... \n" ; 


open( VCF , "<$incfile" )or die("Cannot open vcf file $incfile") ;


#TODO check if header contains required INFO
#Parse VCF header to fill the dictionnary of parameters
print STDERR "Parsing VCF header in order to get sample names and to check if required informations are present ... \n";
my %dicoParam;

my $refGene = 'refGene';
while( <VCF> ){
  	$current_line = $_;
	

	chomp $current_line;
      
    #filling dicoParam with VCF header INFO and FORMAT 

    if ($current_line=~/^##/){

        $vcfHeader .= $_;
      
		  unless ($current_line=~/Description=/){ next }
      #DEBUG print STDERR "Header line\n";

      if ($current_line =~ /ID=(.+?),.*?Description="(.+?)"/){
    

		$dicoParam{$1}= $2;
		if ($1 eq "Func.refGeneWithVer") {$refGene = "refGeneWithVer"}
		#DEBUG      print STDERR "info : ". $1 . "\tdescription: ". $2."\n";
			

			    next;
      
      }else {print STDERR "pattern not found in this line: ".$current_line ."\n";next} 
			
	}elsif($current_line=~/^#CHROM/){
		#check sample names or die
        
        $vcfHeader .= $_;
	
		
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

			#populate non affected-samples array
			if (defined $hashAffected{$line[$sampleIndex]} or (defined $trio and ($line[$sampleIndex] ne $case or $line[$sampleIndex] ne $dad or $line[$sampleIndex] ne $mum))  ){
				#do nothing
			}else{
				push @nonAffectedArray, $line[$sampleIndex]; 
			}
			
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

		#TODO check if all samples in affected list are present or die


		#specify which sample will be the first in the output, to get genotype comment
		# TODO check if it's Ok 
		if (! defined $trio){
				if (@affectedArray){
					$case = $affectedArray[0];
				}elsif($case eq ""){
					$case = $sampleList[0];
				}
		}

		#exclude to treat trio with too much sample
		print STDERR "\nTotal Samples : ".scalar @sampleList."\n";

		if(scalar @sampleList > 3 && defined $trio && scalar @affectedArray == 0 ){
			# TODO check if all affected samples and case , dad and mum are present
			#die("Found more than 3 samples. TRIO analysis is not supported with more than 3 samples.\n");
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
		$workbook = Excel::Writer::XLSX->new( $outDir."/".$outPrefix."achab_catch_newHope.xlsx" );
	}else{
		$workbook = Excel::Writer::XLSX->new( $outDir."/".$outPrefix."achab_catch.xlsx" );
	}
}else {
 	die("No directory $outDir");
}


#create default color background when pLI values are absent
my $format_pLI = $workbook->add_format(bg_color => '#FFFFFF');
my $format_mozaic = $workbook->add_format(bg_color => 'purple');
#$format_pLI -> set_pattern();

#create LOEUF decile associated colors dico

my %dicoLOEUFformatColor;

$dicoLOEUFformatColor{'0.0'} = '#FF0000';
$dicoLOEUFformatColor{'1.0'} = '#FF3300';
$dicoLOEUFformatColor{'2.0'} = '#FF6600';
$dicoLOEUFformatColor{'3.0'} = '#FF9900';
$dicoLOEUFformatColor{'4.0'} = '#FFCC00';
$dicoLOEUFformatColor{'5.0'} = '#FFFF00';
$dicoLOEUFformatColor{'6.0'} = '#BFFF00';
$dicoLOEUFformatColor{'7.0'} = '#7FFF00';
$dicoLOEUFformatColor{'8.0'} = '#3FFF00';
$dicoLOEUFformatColor{'9.0'} = '#00FF00';
$dicoLOEUFformatColor{'.'} = '#FFFFFF';




# Add all worksheets

#$tagsHash{'ALL'} = 'ALL_'.$popFreqThr ; 

my $worksheet = $workbook->add_worksheet('ALL_'.$popFreqThr);
$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLine = 0;

$tagsHash{'ACMG'} = "DS_ACMG" ; 
my $worksheetACMG = $workbook->add_worksheet('DS_ACMG');
$worksheetACMG->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineACMG = 0;

# Meta Data worksheet
$tagsHash{'META'} = "METADATA" ; 
my $worksheetMETA = $workbook->add_worksheet('META');
$worksheetMETA->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineMETA = 0;

# OMIM worksheets
$tagsHash{'OMIMDOM'} = 'OMIM_DOM' ; 
my $worksheetOMIMDOM = $workbook->add_worksheet('OMIM_DOM');
$worksheetOMIMDOM->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineOMIMDOM = 0;

# OMIM worksheets
$tagsHash{'OMIMREC'} = 'OMIM_REC' ; 
my $worksheetOMIMREC = $workbook->add_worksheet('OMIM_REC');
$worksheetOMIMREC->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineOMIMREC = 0;

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

	$tagsHash{'DENOVO'} = 'DENOVO' ; 
	$worksheetDENOVO = $workbook->add_worksheet('DENOVO');
	$worksheetLineDENOVO = 0;
	$worksheetDENOVO->freeze_panes( 1, 0 );    # Freeze the first row

	$tagsHash{'AR'} = 'AR' ; 
	$worksheetAR = $workbook->add_worksheet('AR');
	$worksheetLineAR = 0;
	$worksheetAR->freeze_panes( 1, 0 );    # Freeze the first row
  
	$tagsHash{'HTZ_compo'} = 'HTZ_compo' ; 
	$worksheetHTZcompo = $workbook->add_worksheet('HTZ_compo');
	$worksheetLineHTZcompo = 0;
	$worksheetHTZcompo->freeze_panes( 1, 0 );    # Freeze the first row

	$tagsHash{'SNVmumVsCNVdad'} = 'SNVmumVsCNVdad' ; 
	$worksheetSNPmumVsCNVdad = $workbook->add_worksheet('SNVmumVsCNVdad');
	$worksheetLineSNPmumVsCNVdad = 0;
	$worksheetSNPmumVsCNVdad->freeze_panes( 1, 0 );    # Freeze the first row
  
	$tagsHash{'SNVdadVsCNVmum'} = 'SNVdadVsCNVmum' ; 
	$worksheetSNPdadVsCNVmum = $workbook->add_worksheet('SNVdadVsCNVmum');
	$worksheetLineSNPdadVsCNVmum = 0;
	$worksheetSNPdadVsCNVmum->freeze_panes( 1, 0 );    # Freeze the first row
  

	#$worksheetDELHMZ = $workbook->add_worksheet('DEL_HMZ');
	#$worksheetLineDELHMZ = 0;
	#$worksheetDELHMZ->freeze_panes( 1, 0 );    # Freeze the first row
		

}else{
	
	$tagsHash{'DENOVO'} = 'DOM' ; 
	$worksheetDENOVO = $workbook->add_worksheet('DOM');
	$worksheetLineDENOVO = 0;
	$worksheetDENOVO->freeze_panes( 1, 0 );    # Freeze the first row

	$tagsHash{'AR'} = 'REC' ; 
	$worksheetAR = $workbook->add_worksheet('REC');
	$worksheetLineAR = 0;
	$worksheetAR->freeze_panes( 1, 0 );    # Freeze the first row
	
	if ( @affectedArray){
		$tagsHash{'SNVmumVsCNVdad'} = 'HTZ' ; 
		$worksheetSNPmumVsCNVdad = $workbook->add_worksheet('HTZ');
	
	}else{
		$tagsHash{'SNVmumVsCNVdad'} = 'HMZonly' ; 
		$worksheetSNPmumVsCNVdad = $workbook->add_worksheet('HMZonly');
	}
	$worksheetLineSNPmumVsCNVdad = 0;
	$worksheetSNPmumVsCNVdad->freeze_panes( 1, 0 );    # Freeze the first row

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
			$phenolyzerGene{$current_gene}{'normalized'}= $phenolyzer_List[3];
			$maxLine=0;

		}else{	
			#keep only OMIM or count nbr of lines
			$maxLine+=1;
					
		} 			
	}
	close(PHENO);
} 


if($customVCF_File ne ""){
	
	open(CUSTOMVCF , "<$customVCF_File") or die("Cannot open phenolyzer file ".$customVCF_File) ;
	print  STDERR "Processing custom VCF file ... \n" ; 
	while( <CUSTOMVCF> ){

  		$customVCF_Line = $_;

#############################################
##############   skip header
		next if ($customVCF_Line=~/^#/);

		chomp $customVCF_Line;
		@customVCF_List = split( /\t/, $customVCF_Line );	
		
		#replace "spaces" by "_"
		$customVCF_List[7] =~ s/ /_/g;

		#build variant key as CHROM_POS_REF_ALT

		if (defined $customVCF_variant{$customVCF_List[0]."_".$customVCF_List[1]."_".$customVCF_List[3]."_".$customVCF_List[4]}){
		
			$customVCF_variant{$customVCF_List[0]."_".$customVCF_List[1]."_".$customVCF_List[3]."_".$customVCF_List[4]} .= ";".$customVCF_List[7];

		}else{	

			$customVCF_variant{$customVCF_List[0]."_".$customVCF_List[1]."_".$customVCF_List[3]."_".$customVCF_List[4]} = $customVCF_List[7];
		}
	}
}

#Initialize Regex for filtering customVCF  (should be like that "wantedKey=" )
if ($filterCustomVCFRegex eq ""){
	$filterCustomVCFRegex = "found=";
}else{chomp $filterCustomVCFRegex ;} 


#create sheet for candidats
my $CANDID_TAG = "CANDIDATES";
if($candidates ne ""){
	open( CANDIDATS , "<$candidates")or die("Cannot open candidates file ".$candidates) ;
	print  STDERR "Processing candidates file ... \n" ; 
	while( <CANDIDATS> ){
	  	$candidates_line = $_;
		chomp $candidates_line;

		if($candidates_line =~ m/#/ ){
			$CANDID_TAG = $candidates_line;
			$candidates_line =~ s/#//g;
			$tagsHash{$candidates_line} = $candidates_line; 
	
			next;
		}
		$candidateGene{$candidates_line} .= " ".$CANDID_TAG;

				
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
				$cnvGene{$current_gene} .= " => ".$cnvGene_List[1].";";
				chomp $cnvGene{$current_gene};
		}
			
	}
	close(CNVGENES);

}


#get favourite NM refGene
if($favouriteGeneRef ne ""){
	open( GENEREF , "<$favouriteGeneRef")or die("Cannot open your favourite GeneRef file ".$favouriteGeneRef) ;
	print  STDERR "Processing GeneRef file ... \n" ; 
	while( <GENEREF> ){
	  	$geneRef_line = $_;
		chomp $geneRef_line;
		@geneRefArray = split( /\t/, $geneRef_line);
		$current_gene = $geneRefArray[0];
		if (defined $geneRef_gene{$current_gene}){
				#nothing to do
		}else{	
			$geneRef_gene{$current_gene} = $current_gene;
		}
	}
	close(GENEREF);
}






#Hash of ACMG incidentalome genes
my %ACMGgene = ("ACTA2" =>1,"ACTC1" =>1,"APC" =>1,"APOB" =>1,"ATP7B" =>1,"BMPR1A" =>1,"BRCA1" =>1,"BRCA2" =>1,"CACNA1S" =>1,"COL3A1" =>1,"DSC2" =>1,"DSG2" =>1,"DSP" =>1,"FBN1" =>1,"GLA" =>1,"KCNH2" =>1,"KCNQ1" =>1,"LDLR" =>1,"LMNA" =>1,"MEN1" =>1,"MLH1" =>1,"MSH2" =>1,"MSH6" =>1,"MUTYH" =>1,"MYBPC3" =>1,"MYH11" =>1,"MYH7" =>1,"MYL2" =>1,"MYL3" =>1,"NF2" =>1,"OTC" =>1,"PCSK9" =>1,"PKP2" =>1,"PMS2" =>1,"PRKAG2" =>1,"PTEN" =>1,"RB1" =>1,"RET" =>1,"RYR1" =>1,"RYR2" =>1,"SCN5A" =>1,"SDHAF2" =>1,"SDHB" =>1,"SDHC" =>1,"SDHD" =>1,"SMAD3" =>1,"SMAD4" =>1,"STK11" =>1,"TGFBR1" =>1,"TGFBR2" =>1,"TMEM43" =>1,"TNNI3" =>1,"TNNT2" =>1,"TP53" =>1,"TPM1" =>1,"TSC1" =>1,"TSC2" =>1,"VHL" =>1,"WT1"=>1);





 


#counter for shifting columns according to nbr of sample

my $cmpt = scalar @sampleList;

#Out-of-trio affected samples
my $NTaffectedCmpt = 0;
#Out-of-trio non-affected samples
my $NTnonAffectedCmpt = 0;

#dico for sample sorting (index / dad / mum / control)
my %dicoSamples ;

#dico for mozaic color association to sample
my %hashColor ;

#first: read VCF headers to get refGene or refGeneWithVer
#my $refGene = 'refGene';
#open( VCF , "<$incfile" )or die("Cannot open vcf file ".$incfile) ;
#while( <VCF> ){
#  	$current_line = $_;
#	if ($current_line=~/^##INFO=<ID=Func.refGeneWithVer,/o) {
#		$refGene = 'refGeneWithVer';
#		last;	
#	}
#	elsif ($current_line=~/^##INFO=<ID=Func.refGene,/o) {last}
#}
#
my %dicoColumnNbr;

$dicoColumnNbr{'MPA_ranking'}=				0;	#+ comment MPA scores and related scores 
$dicoColumnNbr{'MPA_impact'}=				1;	#+ comment MPA scores and related scores 
$dicoColumnNbr{'Phenolyzer'}=				2;	#Phenolyzer raw score + comment Normalized score 
$dicoColumnNbr{'Gene.'.$refGene}=				3;  #Gene Name + comment LOEUF / Function_description / tissue specificity
$dicoColumnNbr{'Phenotypes.'.$refGene}=			4;  #OMIM + comment Disease_description
$dicoColumnNbr{'gnomAD_genome_ALL'}=			5;	#Pop Freq + comment ethny
$dicoColumnNbr{'gnomAD_exome_ALL'}=			6;	#as well
$dicoColumnNbr{'CLNSIG'}=				7;	#CLinvar
$dicoColumnNbr{'InterVar_automated'}=			8;	#+ comment ACMG status
$dicoColumnNbr{'SecondHit-CNV'}=			9;	#TODO
$dicoColumnNbr{'Func.'.$refGene}=				10;	# + comment ExonicFunc / AAChange / GeneDetail



if (defined $trio){
	
		$dicoColumnNbr{'Genotype-'.$case}=		11;
		$dicoColumnNbr{'Genotype-'.$dad}=		12;
		$dicoColumnNbr{'Genotype-'.$mum}=		13;

		$dicoSamples{1}{'columnName'} = 'Genotype-'.$case ;
		$dicoSamples{2}{'columnName'} = 'Genotype-'.$dad ;
		$dicoSamples{3}{'columnName'} = 'Genotype-'.$mum ;

		#prepare mozaic colors for each sample at each column position , color is inherit by default 
		$dicoSamples{1}{'columnNbr'} = 11;
		$dicoSamples{2}{'columnNbr'} = 12;
		$dicoSamples{3}{'columnNbr'} = 13;


		# TODO Add out-of-trio samples, affected samples first then non-affected samples 
		 if ( scalar @sampleList > 3){
			if ( scalar  @affectedArray > 0){
				for( my $j = 0 ; $j < scalar @affectedArray; $j++){
					$NTaffectedCmpt++;
					if ($j eq $case or $j eq $dad or $j eq $mum){
						$NTaffectedCmpt--;	
					}else{
						$dicoColumnNbr{'Genotype-'.$affectedArray[$j]}=		13+$NTaffectedCmpt;
						$dicoSamples{3+$NTaffectedCmpt}{'columnName'} = 'Genotype-'.$affectedArray[$j] ;
						$dicoSamples{3+$NTaffectedCmpt}{'columnNbr'} = 13+$NTaffectedCmpt ;

					}
				}
			}

			if ( scalar  @nonAffectedArray > 0){
				for( my $k = 0 ; $k < scalar @nonAffectedArray; $k++){
					$dicoColumnNbr{'Genotype-'.$nonAffectedArray[$k]}=		14+$NTaffectedCmpt+$k;
					$dicoSamples{4+$NTaffectedCmpt+$k}{'columnName'} = 'Genotype-'.$nonAffectedArray[$k] ;
					$dicoSamples{4+$NTaffectedCmpt+$k}{'columnNbr'} = 14+$NTaffectedCmpt+$k ;
					$NTnonAffectedCmpt++;
				}
			}
		}


	
}else{

	#TODO check how many affected samples in this non-trio sample
	
	if ( scalar  @affectedArray > 0){
			for( my $j = 0 ; $j < scalar @affectedArray; $j++){
				$NTaffectedCmpt++;
				$dicoColumnNbr{'Genotype-'.$affectedArray[$j]}=		11+$j;
				$dicoSamples{$j+1}{'columnName'} = 'Genotype-'.$affectedArray[$j] ;
				$dicoSamples{$j+1}{'columnNbr'} = 11+$j ;
			}
			if ( scalar  @nonAffectedArray > 0){
				for( my $k = 0 ; $k < scalar @nonAffectedArray; $k++){
					$dicoColumnNbr{'Genotype-'.$nonAffectedArray[$k]}=		11+$NTaffectedCmpt+$k;
					$dicoSamples{$NTaffectedCmpt+$k+1}{'columnName'} = 'Genotype-'.$nonAffectedArray[$k] ;
					$dicoSamples{$NTaffectedCmpt+$k+1}{'columnNbr'} = 11+$NTaffectedCmpt+$k ;
					$NTnonAffectedCmpt++;
				}
			}
	}else{	
			for( my $i = 0 ; $i < scalar @sampleList; $i++){
	
				$dicoColumnNbr{'Genotype-'.$sampleList[$i]}=			11+$i;	# + comment qual / caller / DP AD AB
				$dicoSamples{$i+1}{'columnName'} = 'Genotype-'.$sampleList[$i] ;
				$dicoSamples{$i+1}{'columnNbr'} = 11+$i ;
			}
	}
}


$dicoColumnNbr{'#CHROM'}=	11+$cmpt ;
$dicoColumnNbr{'POS'}=		12+$cmpt ;
$dicoColumnNbr{'ID'}=		13+$cmpt ;
$dicoColumnNbr{'REF'}=		14+$cmpt ;
$dicoColumnNbr{'ALT'}=		15+$cmpt ;
$dicoColumnNbr{'FILTER'}=	16+$cmpt ;

#Add custom Info in additionnal columns
my $lastColumn = 17+$cmpt; 

if($customInfoList ne ""){
	chomp $customInfoList;
	@custInfList = split(/,/, $customInfoList);
	foreach my $customInfo (@custInfList){
		$dicoColumnNbr{$customInfo} = $lastColumn ;
		$lastColumn++;
	}
}

#add favorite gene references into a new column 
if($favouriteGeneRef ne ""){
	$dicoColumnNbr{'geneRefs'} = $lastColumn ;
	$lastColumn++;
}



#custom VCF in last position
if($customVCF_File ne ""){
	$dicoColumnNbr{'customVCFannotation'} = $lastColumn ;
	$lastColumn++;
}


#TODO add a last column to flag newHope variants
if(defined $newHope){
	$dicoColumnNbr{'Variant_Class'} = $lastColumn ;
	$lastColumn++;
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
my $commentClinvar;

#define sorted arrays with score for comment
my @CommentMPA_score = ("MPA_ranking",
						"MPA_final_score",
						"MPA_impact",
						"MPA_adjusted",
						"MPA_available",
						"MPA_deleterious",
						"\n--- SPLICE ---",
						'dbscSNV_ADA_SCORE',
						'dbscSNV_RF_SCORE',
						'spliceai_filtered',
						#'DS_AG',
						#'DS_AL',
						#'DS_DG',
						#'DS_DL',
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


my @CommentFunc = ( 	'ExonicFunc.'.$refGene,
			'AAChange.'.$refGene,
			'GeneDetail.'.$refGene);


my @CommentPhenotype = ( 'Disease_description.'.$refGene);

my @CommentClinvar = (	'CLNREVSTAT',
			'CLNDN',
			'CLNALLELEID',
			'CLNDISDB');


#For HTML tooltip output
my %commentHash;
$commentHash{'0'}='commentMPAscore';
$commentHash{'2'}='commentPhenolyzer';
$commentHash{'3'}='commentpLI';
$commentHash{'4'}='commentPhenotype';
$commentHash{'5'}='commentGnomADgenome';
$commentHash{'6'}='commentGnomADexome';
$commentHash{'7'}='commentClinvar';
$commentHash{'8'}='commentInterVar';
$commentHash{'10'}='commentFunc';
$commentHash{'11'}='commentGenotype';






#########FILLING COLUMN TITLES FOR SHEETS
$worksheet->write_row( 0, 0, \@columnTitles );
$worksheetACMG->write_row( 0, 0, \@columnTitles );
$worksheetMETA->write( 0, 0, "METADATA" );
$worksheetOMIMDOM->write_row( 0, 0, \@columnTitles ); 
$worksheetOMIMREC->write_row( 0, 0, \@columnTitles ); 

		
#write comment for pLI => NOW LOEUF is used
$worksheet->write_comment( 0, $dicoColumnNbr{'Gene.'.$refGene}, $pLI_Comment,  x_scale => 3 );
$worksheetACMG->write_comment( 0, $dicoColumnNbr{'Gene.'.$refGene}, $pLI_Comment,  x_scale => 3  );


#FILLING COLUMN TITLES FOR TRIO SHEETS OR AFFECTED OR STRANGERS
if (defined $trio){
	$worksheetHTZcompo->write_row( 0, 0, \@columnTitles );
	$worksheetSNPdadVsCNVmum->write_row( 0, 0, \@columnTitles );
			
	#write pLI - LOEUF comment
	$worksheetHTZcompo->write_comment(0,$dicoColumnNbr{'Gene.'.$refGene}, $pLI_Comment,  x_scale => 3);
	$worksheetSNPdadVsCNVmum->write_comment(0, $dicoColumnNbr{'Gene.'.$refGene}, $pLI_Comment,  x_scale => 3);

}

$worksheetAR->write_row( 0, 0, \@columnTitles );
$worksheetDENOVO->write_row( 0, 0, \@columnTitles );
$worksheetSNPmumVsCNVdad->write_row( 0, 0, \@columnTitles );

#write pLI - LOEUF comment
$worksheetAR->write_comment(0, $dicoColumnNbr{'Gene.'.$refGene}, $pLI_Comment,  x_scale => 3);
$worksheetDENOVO->write_comment( 0, $dicoColumnNbr{'Gene.'.$refGene}, $pLI_Comment,  x_scale => 3);
$worksheetSNPmumVsCNVdad->write_comment(0, $dicoColumnNbr{'Gene.'.$refGene}, $pLI_Comment,  x_scale => 3);



#FILLING COLUMN TITLES FOR CANDIDATES SHEET

if($candidates ne ""){
	$worksheetCandidats->write_row( 0, 0, \@columnTitles );
	$worksheetCandidats->write_comment(0, $dicoColumnNbr{'Gene.'.$refGene}, $pLI_Comment,  x_scale => 3 );
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
	$count++;

#############################################
##############   skip header
	next if ($current_line=~/^##/);

	chomp $current_line;
	@line = split( /\t/, $current_line );	
	
	#DEBUG print STDERR $dicoColumnNbr{'Gene.'.$refGene}."\n";



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
	    	$worksheetLineOMIMDOM ++;
	    	$worksheetLineOMIMREC ++;

		if (defined $trio){

			$worksheetLineHTZcompo ++;
			$worksheetLineHTZcompo ++;
			$worksheetLineSNPdadVsCNVmum ++;
		
		}
		$worksheetLineAR ++; 
		$worksheetLineSNPmumVsCNVdad ++;
		$worksheetLineDENOVO ++;

		

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
		$mozaicSamples = "";

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
		#$line[7] =~ s/\\x3d/=/g;
		#$line[7] =~ s/\\x3b/;/g;
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
		next if(( $dicoInfo{'gnomAD_genome_ALL'} ne ".") && ($dicoInfo{'gnomAD_genome_ALL'} > $popFreqThr));  
	
		
		#FILTERING according to newHope option and filterList option
		$filterBool = 0;

		switch ($finalSortData[$dicoColumnNbr{'FILTER'}]){
			case (\@filterArray) {$filterBool=0}
			else {$filterBool=1}
		}
		
		
		if(defined $newHope){
			#Keep only NON PASS or PASS + MPA ranking = 8
			#next if ($finalSortData[$dicoColumnNbr{'FILTER'}] eq "PASS" && $dicoInfo{'MPA_ranking'} < 8);
			#switch ($finalSortData[$dicoColumnNbr{'FILTER'}]){
			#	case (\@filterArray) { $filterBool=1 };
			#}
			#next if($filterBool == 1);
			#next if($dicoInfo{'MPA_ranking'} < 8 );

			if ($filterBool == 1 or $dicoInfo{'MPA_ranking'} == 10){
				$finalSortData[$dicoColumnNbr{'Variant_Class'}] = "New_Hope";	
			}else{
				$finalSortData[$dicoColumnNbr{'Variant_Class'}] = "Main_Catch";	
			}

		}else{
			
			#remove NON PASS variant and remove MPA_Ranking = 10
			next if ($filterBool == 1);
			next if( $dicoInfo{'MPA_ranking'} == 10);
		
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



		#split multiple gene names
		@geneListTemp = split(';', $finalSortData[$dicoColumnNbr{'Gene.'.$refGene}] );	
		
		#uniq genes names
		@geneList = do { my %seen; grep { !$seen{$_}++ } @geneListTemp };	



		#Phenolyzer Column
		if($phenolyzerFile ne ""){

			#@geneList = split(';', $finalSortData[$dicoColumnNbr{'Gene.'.$refGene}] );	

			my $phenoScoreMax=0;	
			foreach my $geneName (@geneList){
				#if (defined $phenolyzerGene{$finalSortData[$dicoColumnNbr{'Gene.'.$refGene}]} )
				if (defined $phenolyzerGene{$geneName} ){
					if( ! defined $finalSortData[$dicoColumnNbr{'Phenolyzer'}]){
						$finalSortData[$dicoColumnNbr{'Phenolyzer'}] = $phenolyzerGene{$geneName}{'Raw'};
					}elsif($phenolyzerGene{$geneName}{'Raw'} > $finalSortData[$dicoColumnNbr{'Phenolyzer'}]){
						$finalSortData[$dicoColumnNbr{'Phenolyzer'}] = $phenolyzerGene{$geneName}{'Raw'};
						
						if($phenoScoreMax < $phenolyzerGene{$geneName}{'normalized'}){
							$phenoScoreMax = $phenolyzerGene{$geneName}{'normalized'};
						}
					}
				
				}

			}

			#refine MPA_rank  with phenolyzer normalized score
			$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += (0.001-($phenoScoreMax/1000 )) ;

		}else{
			$finalSortData[$dicoColumnNbr{'Phenolyzer'}] = ".";
		}

		#CNV GENES
		if($cnvGeneList ne ""){
	
			#@geneList = split(';', $finalSortData[$dicoColumnNbr{'Gene.'.$refGene}] );	
			foreach my $geneName (@geneList){
				#print $geneName."\n";
				if (defined $cnvGene{$geneName} ){
					#print $geneName."toto\n";
					$finalSortData[$dicoColumnNbr{'SecondHit-CNV'}] .= $cnvGene{$geneName}."_";
					#print $finalSortData[$dicoColumnNbr{'SecondHit-CNV'}]."\n";
				}
			}	
		}





		#MPA COMMENT
		#create string with array
		$commentMPAscore = "";
		foreach my $keys (@CommentMPA_score){
			if (defined $dicoInfo{$keys} ){
				if($keys eq "spliceai_filtered"){
					my @dicoSplice = split(';', $dicoInfo{$keys} );
					if (scalar @dicoSplice > 2){
						foreach my $spliceData (@dicoSplice){
							if ($spliceData=~/^DS_/){
								$commentMPAscore .= $keys."\t ".$spliceData."\n";
							}
						}
					}else{
						$commentMPAscore .= $keys."\t= ".$dicoInfo{$keys}."\n";
					}
				}else{	
					$commentMPAscore .= $keys."\t= ".$dicoInfo{$keys}."\n";
				}
			}else{
				$commentMPAscore .= $keys."\n";
			}	
		}

		#refine MPA_rank  with pLI - LOEUF 
		if(($finalSortData[$dicoColumnNbr{'MPA_ranking'}] >= 5 && $finalSortData[$dicoColumnNbr{'MPA_ranking'}] < 6 )|| ($finalSortData[$dicoColumnNbr{'MPA_ranking'}] >= 7 && $finalSortData[$dicoColumnNbr{'MPA_ranking'}] < 8 ) || ($finalSortData[$dicoColumnNbr{'MPA_ranking'}] >= 9 && $finalSortData[$dicoColumnNbr{'MPA_ranking'}] < 10 )  ){
			#$format_pLI = $workbook->add_format(bg_color => $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'});

            #LOEUF refinement





			#if(defined $dicoInfo{'Missense_Z_score.'.$refGene} && $dicoInfo{'Missense_Z_score.'.$refGene} ne "." ){
			if(defined $dicoInfo{'oe_mis_upper.'.$refGene} && $dicoInfo{'oe_mis_upper.'.$refGene} ne "." ){
  				#$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += (0.0001-((($dicoInfo{"Missense_Z_score.".$refGene}+8.64)/22.52)/10000 )) ;

    
    
                #oe_mis_upper  MAX=    MIN=

                $finalSortData[$dicoColumnNbr{'MPA_ranking'}] += $dicoInfo{"oe_mis_upper.".$refGene}/20000 ;


			}else{
				$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += 0.0001;
			}

			#add MPA_final_score before MPA_impact
			$finalSortData[$dicoColumnNbr{'MPA_impact'}] = $dicoInfo{"MPA_deleterious"}."_".$finalSortData[$dicoColumnNbr{'MPA_impact'}];


			#refine MPA_rank for rank 7 missense with MPA final score  
			$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += ((10-$dicoInfo{"MPA_final_score"})/100);
			#print "final_score   :   ".$finalSortData[$dicoColumnNbr{'MPA_ranking'}]."\t".((10-$dicoInfo{"MPA_final_score"})/100)."\t".$dicoInfo{"MPA_final_score"}."\n";
			


		
		}elsif(defined $dicoInfo{"oe_lof_upper.".$refGene} && $dicoInfo{"oe_lof_upper.".$refGene} ne "." ){
		
            #LOEUF  MAX=1.996    MIN=0.03
           
           #clean double LOEUF
            $dicoInfo{"oe_lof_upper.".$refGene} =~ s/;\.//g;
            $finalSortData[$dicoColumnNbr{'MPA_ranking'}] += $dicoInfo{"oe_lof_upper.".$refGene}/20000 ;
            
            
            #$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += (0.0001-(($dicoInfo{"pLi.".$refGene}/10000 +  $dicoInfo{"pRec.".$refGene}/100000 + $dicoInfo{"pNull.".$refGene}/1000000))) ;
					#print $finalSortData[$dicoColumnNbr{'MPA_ranking'}]."\n";
			
			
		}else{
			$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += 0.0001;
		}



		#Penalization of MPA score=7 or 9  to be after at least 9.06 score 
		#no need since MPA v1.1.0

		#if($dicoInfo{"MPA_ranking"} == 7 ){
		#	$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += 2.061;	
		#}elsif( $dicoInfo{"MPA_ranking"} == 9){
		#	$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += 1.062;			
		#}



		#Initialize list of variant tags
		$worksheetTAG = "";

        #Classify OMIM dominant and/or recessive gene variant 
        if (defined $dicoInfo{'Phenotypes.'.$refGene} ){

            if($dicoInfo{'Phenotypes.'.$refGene} =~ m/dominant/ ){ 
                $worksheetTAG .= " OMIMDOM";
            }
            
            if($dicoInfo{'Phenotypes.'.$refGene} =~ m/recessive/ ){ 
                $worksheetTAG .= " OMIMREC";
            }

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
				$dicoInfo{$keys} =~ s/[;,]/\n /g;
				$commentFunc .= $keys.":\n ".$dicoInfo{$keys}."\n\n";
			}
		}

		#TODO check here if transcript reference is present in this comment and extract data in a new column 
		# check among favourite NM references
		if ($favouriteGeneRef ne ""){
			#extract NM_blablah reference and recycling geneRefArray
			if (@geneRefArray = $commentFunc =~ m/(.+(NM_\d+):.+)/g){
				if (@geneRefArray){
					for( my $match = 0 ; $match < scalar @geneRefArray; $match +=2){
						if (defined $geneRef_gene{$geneRefArray[$match+1]}){
							$finalSortData[$dicoColumnNbr{'geneRefs'}] .= $geneRefArray[$match].";";
						}
					}
				}
			}
		}	
		


		#PHENOTYPES REFGENE COMMENT (OMIM)
		#create string with array
		$commentPhenotype = "";
		foreach my $keys (@CommentPhenotype){
			if (defined $dicoInfo{$keys} ){

				$dicoInfo{$keys} =~ s/DISEASE:_/\n\nDISEASE:_/g;
				$commentPhenotype .= $keys.":\n ".$dicoInfo{$keys}."\n\n";
			}
		}


		#CLINVAR COMMENT (CLNSIG)
		#create string with array
		$commentClinvar = "";
		foreach my $keys (@CommentClinvar){
			if (defined $dicoInfo{$keys} ){
				$commentClinvar .= $keys.":\n ".$dicoInfo{$keys}."\n\n";
			}
		}



############################################
#############	Check FORMAT related to caller
		#
		my %formatIndex;
		my @callerFormat = split(':' , $line[8]);
		for (my $z = 0; $z < scalar @callerFormat; $z++){
			$formatIndex{$callerFormat[$z]} = $z;
		}
		

		
		#		GT:AD:DP:GQ:PL (haplotype caller)
		#		GT:DP:GQ  => multiallelic line , vcf not splitted => should be done before, STOP RUN?
		#		GT:GOF:GQ:NR:NV:PL (platyplus caller)
		#		GT:DP:RO:QR:AO:QA:GL (freebayes)
		#		GT:DP:AF (seqNext)
		#
		$caller = "";
		if(defined $formatIndex{'VAF'}){
			$caller = "DeepVariant";
		}elsif(defined $formatIndex{'AD'}){
			$caller = "GATK";
		}elsif(defined $formatIndex{'NR'}){
			$caller = "platypus";
		}elsif(defined $formatIndex{'AO'}){
			$caller = "freebayes";
		}elsif(defined $formatIndex{'AF'} ){ 
			# "GT:DP:AF"
			#SeqNext like format
			$caller = "other _ GT:DP:AF";

		}elsif(defined $formatIndex{'DP'}){
			$caller = "other _ GT:DP";
			#print STDERR "Multi-allelic line detection. Please split the vcf file in order to get 1 allele by line\n";
			#print STDERR $current_line ."\n";
			#exit 1; 



		}else{
			print STDERR "The Format of the Caller used for this line is unknown. Processing is a risky business. This line won't be processed.\n";
			print STDERR $current_line ."\n";
			$caller = "unknown";
			next;
			#exit 1; 
		}
		



		#genotype concatenation for easy hereditary status
		$familyGenotype = "_";		
		$commentGenotype = "CALLER = ".$caller."\t QUALITY = ".$line[5]."\n\n";

		#DEBUG print STDERR "commentgeno     :     ".$commentGenotype."\n";

#############################################################################
#########FILL HASH STRUCTURE FOR FINAL SORT AND OUTPUT, according to rank
		
		#concatenate chrom_POS_REF_ALT to get variant ID
		$variantID = $line[0]."_".$line[1]."_".$line[3]."_".$line[4]."_".$count;
		#$variantID = $line[0]."_".$line[1]."_".$line[3]."_".$line[4]."_".$caller;
		

		#CHECK IF variant is in customVCF
		if($customVCF_File ne "" && defined $customVCF_variant{$line[0]."_".$line[1]."_".$line[3]."_".$line[4]}){
			$finalSortData[$dicoColumnNbr{'customVCFannotation'}] = $customVCF_variant{$line[0]."_".$line[1]."_".$line[3]."_".$line[4]};

			#Penalize variant if "found=*" is greater than filterCustomVCF option
			if($filterCustomVCF ne "" && $finalSortData[$dicoColumnNbr{'customVCFannotation'}] =~ m/$filterCustomVCFRegex(\d*)/){
				if( $1 >= $filterCustomVCF){ 
					$finalSortData[$dicoColumnNbr{'MPA_ranking'}]   += 100;
				}
			}
		}



############################################
#############   Parse Genotypes
		#

		#for each sample sort by sample wanted 
		foreach my $finalcol ( sort {$a <=> $b}  (keys %dicoSamples) ) {

			#DEBUG print "tata\t".$finalcol."\n";
			#DEBUG	print $line[$dicoSamples{$finalcol}{'columnIndex'}]."\n";
				
				my @genotype = split(':', $line[$dicoSamples{$finalcol}{'columnIndex'}] );
				#my $genotype =  $line[$dicoSamples{$finalcol}{'columnIndex'}];
				
                my $DP; #total Depth
    
                if (defined $formatIndex{'DP'}){
				    $DP = $genotype[$formatIndex{'DP'}]; 
                }


				my $adalt;	#Alternative Allelic Depth
				my $adref;	#Reference Allelic Depth
				my $AB;		#Allelic balancy
				my $AD;		#Final Allelic Depth


				#if (scalar @genotype > 1 && $caller ne ""){
				if ($caller ne "unknown"){

					if(	($caller eq "GATK" ) || ($caller eq "DeepVariant")){

						#check if variant is not called
						if ($DP eq "."){
							$DP = 0;
							$AD = "0,0";
						
						}else{
							$AD = $genotype[$formatIndex{'AD'}];
						}

					}elsif($caller eq "platypus"){
						
						if($genotype[$formatIndex{'NR'}] =~ m/,/){
							my @genotype_DPsplit = split(',', $genotype[$formatIndex{'NR'}]);
							my @genotype_ADsplit = split(',', $genotype[$formatIndex{'NV'}]);
							$DP = $genotype_DPsplit[0];
							$AD = ($genotype_DPsplit[0] - $genotype_ADsplit[0]);
							$AD .= ",".$genotype_ADsplit[0];
							
						}else{
							$DP = $genotype[$formatIndex{'NR'}];
							$AD = ($genotype[$formatIndex{'NR'}] - $genotype[$formatIndex{'NV'}]);
							$AD .= ",".$genotype[$formatIndex{'NV'}];
						}
					}elsif($caller eq "other _ GT:DP"){
						
						$AD = 0;
						$AD .= ",0";
						
					}elsif($caller eq "other _ GT:DP:AF"){
						$AD = ($DP-int($DP*$genotype[$formatIndex{'AF'}])) . "," . int($DP*$genotype[$formatIndex{'AF'}]);
					
					}elsif($caller eq "freebayes"){
						$AD = $genotype[$formatIndex{'RO'}];
						$AD .= ",".$genotype[$formatIndex{'AO'}];
					}
					
					#DEBUG
					#print $genotype[2]."\n";


					#split allelic depth
					my @tabAD = split( ',',$AD);
					
					#print "\ntabAD\t",split( ',',$AD),"\n";

					if ( scalar @tabAD > 1 ){
						$adref = $tabAD[0];
						$adalt = $tabAD[1];
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


				}else {  # caller unknown           @genotype length < 1
					$adref = 0;
					$adalt= 0;
					$AB = 0;
					$DP = $line[8];
					$AD = $line[$dicoSamples{$finalcol}{'columnIndex'}];
				}


				#DEBUG print STDERR "indexSample\t".$dicoSamples{$finalcol}{'columnIndex'}."\n";

				#put the genotype and comments info into string
				#convert "1/0" genotype to "0/1" format
				if($genotype[$formatIndex{'GT'}] eq "1/0"){
					$genotype[$formatIndex{'GT'}] = "0/1";
				}
				$finalSortData[$dicoColumnNbr{$dicoSamples{$finalcol}{'columnName'}}] = $genotype[$formatIndex{'GT'}];
				$commentGenotype .=  $dicoSamples{$finalcol}{'columnName'}."\t -\t ".$genotype[$formatIndex{'GT'}]."\nDP = ".$DP."\t AD = ".$AD."\t AB = ".$AB."\n\n";
				$familyGenotype .= $genotype[$formatIndex{'GT'}]."_";
				
				#record mozaic status of samples and (TODO) low cov for 1/1 genotypes 
				if(($genotype[$formatIndex{'GT'}] eq "0/1" && $AB < $mozaicRate) or ($genotype[$formatIndex{'GT'}] eq "1/1" && $adalt < $mozaicDP     )){
					
					$mozaicSamples  .= $dicoSamples{$finalcol}{'columnName'}.";";

					#penalty to MPA ranking if trio and mozaic for the index case 
					if(defined $trio && $dicoSamples{$finalcol}{'columnName'} eq "Genotype-".$case){
						$finalSortData[$dicoColumnNbr{'MPA_ranking'}] 	+= 0.1;

						#penalty to low covered ALT base
						if($adalt < $mozaicDP){
							$finalSortData[$dicoColumnNbr{'MPA_ranking'}]   += 0.1;
						}				
					}
					
					if($genotype[$formatIndex{'GT'}] eq "1/1"){
						$mozaicSamples  .= 'cyan'.";";
						$hashColor{$dicoSamples{$finalcol}{'columnName'}} = 'cyan';

					}elsif($adalt < $mozaicDP){
						$mozaicSamples  .= 'purple'.";";
						$hashColor{$dicoSamples{$finalcol}{'columnName'}} = 'purple';
					}else{				
						$mozaicSamples  .= 'pink'.";";
						$hashColor{$dicoSamples{$finalcol}{'columnName'}} = 'pink';
					}
					

				}else{
					$hashColor{$dicoSamples{$finalcol}{'columnName'}} = 'inherit';
				}


		} #END of Sample Treatment	


###################################################################################
######################### GENOTYPE ANALYSIS #######################################
###################################################################################
########## additionnal analysis in TRIO or affected context according to family genotype + CNV


		#Penalize (or do next) if index case is 0/0 or parents are 1/1 and not affected. We should treat further all affected genotypes like this (!= 0/0)
		if (defined $trio){ 
			switch ($familyGenotype){
				#Check if case/dad and case/mum  inheritance are  consistent
				case /^_0\/0_0\/1_0\/0_/ {$dadVariant ++ ;}
				case /^_0\/0_0\/0_0\/1_/ {$mumVariant ++ ;}
				case /^_0\/1_0\/1_0\/0_/ {$caseDadVariant ++;}
				case /^_0\/1_0\/0_0\/1_/ {$caseMumVariant ++;}

			}

			if ($finalSortData[$dicoColumnNbr{"Genotype-".$case}] eq "0/0" or (! defined $hashAffected{$dad} and $finalSortData[$dicoColumnNbr{"Genotype-".$dad}] eq "1/1") or (! defined $hashAffected{$mum} and $finalSortData[$dicoColumnNbr{"Genotype-".$mum}] eq "1/1") ){
				$finalSortData[$dicoColumnNbr{'MPA_ranking'}]   += 100;
			}
		}elsif (@affectedArray){
			foreach my $AFF (@affectedArray){
				if ($finalSortData[$dicoColumnNbr{"Genotype-".$AFF}] eq "0/0"){
					$finalSortData[$dicoColumnNbr{'MPA_ranking'}]   += 100;
					last;
				}
			}	
			if ( scalar  @nonAffectedArray > 0){
				if ($finalSortData[$dicoColumnNbr{'MPA_ranking'}]   < 10){;
					foreach my $NAFF (@nonAffectedArray){
						if ($finalSortData[$dicoColumnNbr{"Genotype-".$NAFF}] eq "1/1"){
							$finalSortData[$dicoColumnNbr{'MPA_ranking'}]   += 100;
							last;
						}
					}
				}
			}
		}
		# TODO, add elsif with a foreach loop with affected that shouldn't be 0/0 and non-affected 1/1
	
		#

		if (defined $trio){
			
			switch ($familyGenotype){		#INFO you must use this switch syntax: case m/myRegex/ with complex regex (instead of case /regex/

				#find Autosomique Recessive
				#case ["_1/1_0/1_0/1_"] {$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#AUTOREC";}
				#TODO
				case m/^_1\/1_0\/1_0\/1_(1\/1_){$NTaffectedCmpt}(0\/1_){$NTnonAffectedCmpt}/ {$worksheetTAG .= " AUTOREC";}


				#Find de novo
				#case ["_1/1_0/0_0/0_" ,"_0/1_0/0_0/0_"] {$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#DENOVO";
					
				#TODO 
				case m/^_[01]\/1_0\/0_0\/0_([01]\/1_){$NTaffectedCmpt}(0\/0_){$NTnonAffectedCmpt}/ {$worksheetTAG .= " DENOVO";

					foreach my $geneName (@geneList){
						unless(defined $dicoGeneForHTZcompo{$geneName}{'denovo'} ){$dicoGeneForHTZcompo{$geneName}{'denovo'} = 1 ;}
						$dicoGeneForHTZcompo{$geneName}{'variantID'} .= $variantID."#";
						$dicoGeneForHTZcompo{$geneName}{'MPA_ranking'} .= $finalSortData[$dicoColumnNbr{'MPA_ranking'}]."#" ;
					}
				}




				#Find HTZ composite
				#case ["_0/1_0/1_0/0_"] {
				#TODO 
				case m/^_0\/1_0\/1_0\/0_(0\/1_){$NTaffectedCmpt}(0\/[01]_){$NTnonAffectedCmpt}/ {
					
					foreach my $geneName (@geneList){
						unless(defined $dicoGeneForHTZcompo{$geneName}{'pvsm'} ){$dicoGeneForHTZcompo{$geneName}{'pvsm'} = 1 ;	}

						$dicoGeneForHTZcompo{$geneName}{'variantID'} .= $variantID."#";
						$dicoGeneForHTZcompo{$geneName}{'MPA_ranking'} .= $finalSortData[$dicoColumnNbr{'MPA_ranking'}]."#" ;
					}
				}	
				
				#case ["_0/1_0/0_0/1_"] { 	
				#TODO 
				case m/^_0\/1_0\/0_0\/1_(0\/1_){$NTaffectedCmpt}(0\/[01]_){$NTnonAffectedCmpt}/ {

					foreach my $geneName (@geneList){
						unless(defined $dicoGeneForHTZcompo{$geneName}{'mvsp'} ){$dicoGeneForHTZcompo{$geneName}{'mvsp'} = 1 ;}

						$dicoGeneForHTZcompo{$geneName}{'variantID'} .= $variantID."#";
						$dicoGeneForHTZcompo{$geneName}{'MPA_ranking'} .= $finalSortData[$dicoColumnNbr{'MPA_ranking'}]."#" ;
					}
				}
				
				#Find SNVvsCNV
				#case ["_1/1_0/0_0/1_"] {$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#SNPmCNVp";}
				#TODO 
				case m/^_1\/1_0\/0_0\/1_(1\/1_){$NTaffectedCmpt}(0\/[01]_){$NTnonAffectedCmpt}/ {$worksheetTAG .= " SNPmCNVp";}

				#case ["_1/1_0/1_0/0_"] {$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#SNPpCNVm";}
				#TODO 
				case m/^_1\/1_0\/1_0\/0_(1\/1_){$NTaffectedCmpt}(0\/[01]_){$NTnonAffectedCmpt}/ {$worksheetTAG  .= " SNPpCNVm";}

					
				#case ["_0/1_0/1_0/1_"]	{
				#TODO 
				case m/^_0\/1_0\/1_0\/1_(0\/1_){$NTaffectedCmpt}(0\/1_){$NTnonAffectedCmpt}/ {
					foreach my $geneName (@geneList){
						if(defined $dicoGeneForHTZcompo{$geneName}{'any'} ){
#							$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#SNPpCNVm";}
							# do nothing						
						}else{
							$dicoGeneForHTZcompo{$geneName}{'any'} = 100 ;
						}
						$dicoGeneForHTZcompo{$geneName}{'variantID_any'} .= $variantID."#"; 
						$dicoGeneForHTZcompo{$geneName}{'MPA_ranking_any'} .= $finalSortData[$dicoColumnNbr{'MPA_ranking'}]."#";
					}
				}		
									
			}#END SWITCH 
			
			#Check if CNV is present
			if (defined $finalSortData[$dicoColumnNbr{'SecondHit-CNV'}] && $finalSortData[$dicoColumnNbr{'SecondHit-CNV'}] ne "."){
				foreach my $geneName (@geneList){
					if(defined $dicoGeneForHTZcompo{$geneName}{'CNV'} ){
#						$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#SNPpCNVm";}
						# do nothing						
					}else{
							$dicoGeneForHTZcompo{$geneName}{'CNV'} = 1 ;
					}

						$dicoGeneForHTZcompo{$geneName}{'variantID'} .= $variantID."#";
						$dicoGeneForHTZcompo{$geneName}{'MPA_ranking'} .= $finalSortData[$dicoColumnNbr{'MPA_ranking'}]."#" ;
				}
					
			}

		}elsif (scalar @affectedArray > 0){						#END IF TRIO
			
			# TODO open additionnal tab in non-trio contexte (affected and stangers)
			#Family Affected mode (non-trio)
			switch ($familyGenotype){
				
				#find Autosomique Recessive
				case m/^_(1\/1_){$NTaffectedCmpt}(0\/1_){$NTnonAffectedCmpt}/ {$worksheetTAG .= " AUTOREC";}

				#Find de novo
				case m/^_(1\/1_){$NTaffectedCmpt}(0\/0_){$NTnonAffectedCmpt}/ {$worksheetTAG .= " DENOVO";}
				case m/^_(0\/1_){$NTaffectedCmpt}(0\/0_){$NTnonAffectedCmpt}/ {$worksheetTAG .= " DENOVO";}

				#Find HTZ composite or should be "dominant" term
				case m/^_(0\/1_){$NTaffectedCmpt}(0\/[01]_){$NTnonAffectedCmpt}/ {$worksheetTAG .= " SNPmCNVp";}

			}
				

		}else{  #END OF AFFECTED
			
			#Stranger Mode
			@strangerNULL = $familyGenotype =~ m/\.\/\./g;
			@strangerREF = $familyGenotype =~ m/0\/0/g;
			@strangerHTZ = $familyGenotype =~ m/0\/1/g;
			@strangerHMZ = $familyGenotype =~ m/1\/1/g;

			if (scalar @strangerHTZ > 0 && scalar @strangerHTZ < 2 && (scalar @strangerNULL + scalar @strangerREF + scalar @strangerHTZ == $cmpt)){
				$worksheetTAG .= " DENOVO";
				#$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#DENOVO";
			}elsif (scalar @strangerHMZ > 0 && scalar @strangerHMZ < 2 && (scalar @strangerNULL + scalar @strangerREF + scalar @strangerHTZ + scalar @strangerHMZ == $cmpt)){
				$worksheetTAG .= " AUTOREC";
			}elsif (scalar @strangerHMZ > 0 && scalar @strangerHMZ < 2 && (scalar @strangerNULL + scalar @strangerREF + scalar @strangerHMZ == $cmpt)){
				$worksheetTAG .= " SNPmCNVp";
			}

		}  # END OF STRANGER MODE




		
		#	print Dumper(\@finalSortData);

#############################################################################
######### FILL HASH STRUCTURE WITH COMMENTS 
		
		
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'finalArray'} = [@finalSortData] ; 
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'nbSample'} = scalar keys %dicoSamples ;
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentGnomADexome'} = $commentGnomADExomeScore  ;
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentGnomADgenome'} = $commentGnomADGenomeScore ;
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentGenotype'} = $commentGenotype ;
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentMPAscore'} = $commentMPAscore  ;
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentInterVar'} = $commentInterVar  ;
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentFunc'} = $commentFunc  ;
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentPhenotype'} = $commentPhenotype  ;
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentClinvar'} = $commentClinvar  ;
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'genotypeMozaic'} = $mozaicSamples ;

		#initialize worksheet
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} = $worksheetTAG;
		#$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} = "";

		#test multiple geneNames
		#@geneList = split(';', $finalSortData[$dicoColumnNbr{'Gene.'.$refGene}] );	
			
		foreach my $geneName (@geneList){

			#ACMG
			#if(defined $ACMGgene{$finalSortData[$dicoColumnNbr{'Gene.'.$refGene}]} )
			if(defined $ACMGgene{$geneName} ){
				$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= " ACMG";
			}
				
			#CANDIDATES
			if($candidates ne ""){
				#if(defined $candidateGene{$finalSortData[$dicoColumnNbr{'Gene.'.$refGene}]} )
				if(defined $candidateGene{$geneName} ){
					#$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#CANDIDATES";
					$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .=    " ".$candidateGene{$geneName};		
				}
				
			}

			#PHENOLYZER COMMENT AND SCORE
				
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentPhenolyzer'} = "";
				
			#if(defined  $phenolyzerGene{$finalSortData[$dicoColumnNbr{'Gene.'.$refGene}]})
			if(defined  $phenolyzerGene{$geneName}){
				$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentPhenolyzer'} .= $phenolyzerGene{$geneName}{'comment'}."\n\n"  ;
					
			}
		
		}#END FOREACH




###############################################################
		# TODO switch to LOEUF system
##########create pLI comment and format
		#if(defined $dicoInfo{'pLi.'.$refGene} && $dicoInfo{'pLi.'.$refGene} ne "." ){
		if(defined $dicoInfo{'oe_lof_upper_bin.'.$refGene} && $dicoInfo{'oe_lof_upper_bin.'.$refGene} ne "." ){

			#$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} =  "pLI = ".$dicoInfo{'pLi.'.$refGene}."\npRec = ".$dicoInfo{'pRec.'.$refGene}."\npNull = ".$dicoInfo{'pNull.'.$refGene} ."\n\n";

#            oe_lof_upper_rank   oe_lof_upper_bin    oe_lof  oe_lof_lower    oe_lof_upper    oe_mis  oe_mis_lower    oe_mis_upper

			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} =  "LOEUF = ".$dicoInfo{'oe_lof_upper.'.$refGene}."\nLOEUF_decile = ".$dicoInfo{'oe_lof_upper_bin.'.$refGene}."\n\noe_lof = ".$dicoInfo{'oe_lof.'.$refGene}."\noe_lof_lower = ".$dicoInfo{'oe_lof_lower.'.$refGene} ."\n\n";

            if( $dicoInfo{'oe_lof_upper_bin.'.$refGene} =~ /;/){
                
                my @tweenBin = split( ';', $dicoInfo{'oe_lof_upper_bin.'.$refGene}   );    

                if ($tweenBin[1] eq "." && $tweenBin[0] ne "."){
			        $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  $dicoLOEUFformatColor{$tweenBin[0]} ;
                }elsif($tweenBin[0] eq "." && $tweenBin[1] ne "."){
                    
			        $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  $dicoLOEUFformatColor{$tweenBin[1]} ;
                }elsif($tweenBin[0] eq "." && $tweenBin[1] eq "."){
                    $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  '#FFFFFF' ;
                
                }elsif($tweenBin[0] >= $tweenBin[1]){
                     $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  $dicoLOEUFformatColor{$tweenBin[1]} ;
                }else{
                     $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  $dicoLOEUFformatColor{$tweenBin[0]} ;    
                }

            }else{

			#$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  sprintf('#%2.2X%2.2X%2.2X',($dicoInfo{'pLi.'.$refGene}*255 + $dicoInfo{'pRec.'.$refGene}*255),($dicoInfo{'pRec.'.$refGene}*255 + $dicoInfo{'pNull.'.$refGene} * 255),0) ;
        		
			    $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  $dicoLOEUFformatColor{$dicoInfo{'oe_lof_upper_bin.'.$refGene}} ;

            }
            
            #add color format
			$format_pLI = $workbook->add_format(bg_color => $hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'});



			if(defined $dicoInfo{'oe_mis_upper.'.$refGene} && $dicoInfo{'oe_mis_upper.'.$refGene} ne "." ){
			#if(defined $dicoInfo{'Missense_Z_score.'.$refGene} && $dicoInfo{'Missense_Z_score.'.$refGene} ne "." ){
				#$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} .=  "Missense Z-score = ".$dicoInfo{'Missense_Z_score.'.$refGene}."\n\n";
				$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} .=  "Missense_upper = ".$dicoInfo{'oe_mis_upper.'.$refGene}."\n\n";
			}
		
		}else{	
			
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} =  "." ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  '#FFFFFF' ;
			
			#$format_pLI = $workbook->add_format(bg_color => '#FFFFFF');
		}

###############################################################
##########Add function and expression infos in comment
		if(defined $dicoInfo{'Function_description.'.$refGene}  && $dicoInfo{'Function_description.'.$refGene} ne "." ){
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} .= "Function Description:\n".$dicoInfo{'Function_description.'.$refGene}."\n\n"; 
			
			if(defined $dicoInfo{'Tissue_specificity(Uniprot).'.$refGene} && $dicoInfo{'Tissue_specificity(Uniprot).'.$refGene} ne "."  ){
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} .= "Tissue specificity:\n".$dicoInfo{'Tissue_specificity(Uniprot).'.$refGene}."\n"; 	
			} 

		} 
		

	} #END of IF-ELSE(#CHROM)	





#DEBUG			print Dumper(\%dicoColumnNbr);
##############check hereditary hypothesis or genes to fill sheets





############ THE TIME HAS COME TO FILL THE XLSX OUTPUT FILE





}#END WHILE VCF


########################################################################
###################	Compute HTZ compo before writing

foreach my $geneName (sort keys %dicoGeneForHTZcompo){
		

	#my $MPA = $dicoGeneForHTZcompo{$geneName}{'HTZscore'};
	#my $variant = $dicoGeneForHTZcompo{$geneName}{'HTZscore'};

	if(defined 	$dicoGeneForHTZcompo{$geneName}{'CNV'}){
		$dicoGeneForHTZcompo{$geneName}{'HTZscore'} += 1  ;
	}
	if(defined 	$dicoGeneForHTZcompo{$geneName}{'denovo'}){
		$dicoGeneForHTZcompo{$geneName}{'HTZscore'} += 1  ;
	}
	if(defined 	$dicoGeneForHTZcompo{$geneName}{'HTZscore'} && $dicoGeneForHTZcompo{$geneName}{'HTZscore'} > 0 && defined $dicoGeneForHTZcompo{$geneName}{'any'} ){
		$dicoGeneForHTZcompo{$geneName}{'HTZscore'} += 100;  
	}
	if(defined 	$dicoGeneForHTZcompo{$geneName}{'pvsm'}){
		$dicoGeneForHTZcompo{$geneName}{'HTZscore'} += 1 ;
	}
	if(defined 	$dicoGeneForHTZcompo{$geneName}{'mvsp'}){
		$dicoGeneForHTZcompo{$geneName}{'HTZscore'} += 1  ;
	}
	
	if(defined $dicoGeneForHTZcompo{$geneName}{'HTZscore'}  && $dicoGeneForHTZcompo{$geneName}{'HTZscore'} > 1  ){
		#DEBUG
		#print $geneName."\t".$dicoGeneForHTZcompo{$geneName}{'HTZscore'}."\n";
		
		
		my @variantList = split( '#', $dicoGeneForHTZcompo{$geneName}{'variantID'} );
		my @MPAList = split( '#', $dicoGeneForHTZcompo{$geneName}{'MPA_ranking'} );
		
		for(my $i=0; $i < scalar @variantList; $i ++){
			
			unless(defined	$hashFinalSortData{$MPAList[$i]}{$variantList[$i]}{'worksheetHTZcompo'}){
				$hashFinalSortData{$MPAList[$i]}{$variantList[$i]}{'worksheetHTZcompo'} = "#HTZcompo";	
				

				#create reference of Hashes
				my $hashTemp = $hashFinalSortData{$MPAList[$i]}{$variantList[$i]};
				my $hashColumn_ref = \%dicoColumnNbr;
			
##############################################################
###########################      HTZ composite     #####################

				#Write the HTZcompo worksheet
				writeThisSheet ($worksheetHTZcompo,
								$worksheetLineHTZcompo,
								$format_pLI,
								$case,
								$hashTemp,
								$hashColumn_ref
								);
				$worksheetLineHTZcompo ++;

			}#END unless 

		}#END FOR



	}#END HTZscore >1

	if( defined $dicoGeneForHTZcompo{$geneName}{'HTZscore'} && $dicoGeneForHTZcompo{$geneName}{'HTZscore'} > 100  ){
		#DEBUG
		#print $geneName."\t".$dicoGeneForHTZcompo{$geneName}{'HTZscore'}."\n";
		
		
		my @variantList = split( '#', $dicoGeneForHTZcompo{$geneName}{'variantID_any'} );
		my @MPAList = split( '#', $dicoGeneForHTZcompo{$geneName}{'MPA_ranking_any'} );
		
		for(my $i=0; $i < scalar @variantList; $i ++){
			
			unless(defined	$hashFinalSortData{$MPAList[$i]}{$variantList[$i]}{'worksheetHTZcompo'}){
				$hashFinalSortData{$MPAList[$i]}{$variantList[$i]}{'worksheetHTZcompo'} = "#HTZcompo";	
				

				#create reference of Hashes
				my $hashTemp = $hashFinalSortData{$MPAList[$i]}{$variantList[$i]};
				my $hashColumn_ref = \%dicoColumnNbr;
			
##############################################################
###########################      HTZ composite     #####################

				#Write the HTZcompo worksheet
				writeThisSheet ($worksheetHTZcompo,
								$worksheetLineHTZcompo,
								$format_pLI,
								$case,
								$hashTemp,
								$hashColumn_ref
								);
				$worksheetLineHTZcompo ++;

			}#END unless 

		}#END FOR

	}#END HTZscore >100

}

####################################################################
######################## HTML  Initialisation ######################

#file header




my $htmlStart = "<!DOCTYPE html>\n<html>
\n<head>
\n<meta charset=\"utf-8\">
\n<title>".$outPrefix." Achab catch</title>\n
\n<script type=\"text/javascript\" language=\"javascript\" src='https://code.jquery.com/jquery-3.5.1.js'></script>
\n<script type=\"text/javascript\" language=\"javascript\" src='https://cdn.datatables.net/fixedheader/3.1.7/js/dataTables.fixedHeader.min.js'></script>
\n<script type=\"text/javascript\" language=\"javascript\" src='https://cdn.datatables.net/1.10.22/js/jquery.dataTables.min.js'></script>
\n<script>

var filter;

\$(document).ready(function () {


	\$('#table thead tr').clone(true).appendTo( '#table thead' );
            \$('#table thead tr:eq(1) th').each( function (i) {
             var title = \$(this).text();
              \$(this).html( '<input type=\"text\" placeholder=\"Search\" />' );
                                                 
              \$( 'input', this ).on( 'keyup change', function () {
                  if ( table.column(i).search() !== this.value ) {
                        table
                       .column(i)
                       .search( this.value,true,false )
                       .draw();
                  }
            } );
        } );
	



	var table = \$('#table').DataTable(        {\"order\": [] ,\"lengthMenu\":[ [ 50, 100, -1 ],[ 50, 100, \"All\" ]], \"fixedHeader\": true, \"orderCellsTop\": true, \"oLanguage\": { \"sLengthMenu\": \"Show _MENU_ lines\",\"sInfo\": \"Showing _START_ to _END_ of _TOTAL_ lines\" } } );


filter = function  (cat) {

	var rowunselected = document.getElementsByTagName(\"TR\");
	for (i = 0; i < rowunselected.length; i++) {
		rowunselected[i].style.visibility = \"collapse\";
	}
	var rowselect = document.getElementById('table').getElementsByClassName(cat);
	for (i = 0; i < rowselect.length; i++) {
		rowselect[i].style.visibility = \"visible\";
	}
	     
	var rowhead = document.getElementById('table').getElementsByClassName('head');
	for (i = 0; i < rowhead.length; i++) {
		rowhead[i].style.visibility = \"visible\";
	}
}




});


</script>
\n<link rel=\"stylesheet\" type=\"text/css\" href='https://cdn.datatables.net/1.10.22/css/jquery.dataTables.min.css'>
\n<link rel=\"stylesheet\" type=\"text/css\" href='https://cdn.datatables.net/fixedheader/3.1.7/css/fixedHeader.dataTables.min.css'>

\n<style>



/* Style the tab */
        .tab {
                overflow: hidden;
                border: 1px solid #ccc;
                background-color: #f1f1f1;
                left: 0;
                bottom: 0;
                position: fixed;
                position: -webkit-sticky;
                position: sticky;
                }

/* Style the buttons inside the tab */
        .tab input {
                background-color: inherit;
                float: left;
                border: none;
                outline: none;
                cursor: pointer;
                padding: 14px 16px;
                transition: 0.3s;
                font-size: 17px;
                }
/* Change background color of buttons on hover */
        .tab input:hover {
                background-color: #ddd;
                }

/* Create an active/current tablink class */
        .tab input.active {
                background-color: #ccc;
                }
        .tab input:focus {
                background-color: #ccc;
                }

        thead input {
                width: 100%;
                }


/* Style the tab content */
        .tabcontent {
                visibility: visible;
                padding: 6px 12px;
                /*border: 1px solid #ccc;*/
                border-top: none;
                }

/* Style the tab content FULL */

        .tooltip {
                position: relative;
                display: inline-block;
                border-bottom: 1px dotted black;
                max-width: 300px;
                word-wrap: break-word;
                }
        
        .tooltip .tooltiptext {
                visibility: hidden;
                width: auto;
                min-width: 300px;
                max-width: 600px;
                height: auto;
                background-color: #555;
                color: #fff;
                text-align: left;
                border-radius: 6px;
                padding: 5px 0;
                position: absolute;
                z-index: 1;
                top: 100%;
                left: 20%;
                margin-left: -20px;
                opacity: 0;
                transition: opacity 0.3s;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: pre-line;
                overflow-wrap: break-word;
                }

        .tooltip:hover .tooltiptext {
                visibility: visible;
                opacity: 1;
                }



\n</style>



\n</head>
\n\t<body>";

#table and columns names



my $htmlStartTable = "";
foreach my $col (@columnTitles){
	#print HTML "\t<th style=\"word-wrap: break-word\"   >";
	$htmlStartTable .= "\t<th >".$col."\t</th>\n";
}
$htmlStartTable .= "</tr>\n</thead>\n<tbody>\n";



my $htmlALL= "<div id=\"ALL\" class=\"tabcontent\">\n\t<table id='table' class=\"display compact\" >\n\t\t<thead><tr class=\"head\">".$htmlStartTable;
 
#my $htmlMETA="<div id=\"METADATA\" class=\"tabcontent\">\n\t<table id='tabMETADATA' class='display' >\n\t\t<thead><tr>".$htmlStartTable;




my $htmlEndTable = "</tbody>\n</table>\n</div>\n\n\n";

my $htmlEnd = "<div class=\"tab\">\n<input type=\"button\" class=\"tablinks\" value=\"ALL\" onclick=\"filter('ALL')\">\n" ;
	foreach my $tag ( keys %tagsHash){
		$htmlEnd .= "<input type=\"button\" class=\"tablinks\" value=\"".$tagsHash{$tag}."\" onclick=\"filter('".$tag."')\" >\n" ;
	}
$htmlEnd .= "</div>" ;

$htmlEnd .= "\n</body>\n</html>";





open(HTML, '>', $outDir."/".$outPrefix."achab.html") or die $!;


#########################################################################
#################### Sort by MPA ranking for the output

#create user friendly ranking score
my $kindRank=0;


foreach my $rank (sort {$a <=> $b} keys %hashFinalSortData){
	#print $rank."\n";
	
	foreach my $variant ( keys %{$hashFinalSortData{$rank}}){

#		$format_pLI = $workbook->add_format(bg_color => '#FFFFFF');
		
		#increse rank number then change final array
		$kindRank++;
		
		#print $variant ."___".  $hashFinalSortData{$rank}{$variant}{'finalArray'}[$dicoColumnNbr{'MPA_ranking'}]."\n";
		$hashFinalSortData{$rank}{$variant}{'finalArray'}[$dicoColumnNbr{'MPA_ranking'}] = $kindRank; 

		#last finalSortData assignation
		$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'finalArray'} = [@finalSortData] ; 
		
		#print Dumper($hashFinalSortData{$rank}{$variant}{'finalArray'});

		#create reference of Hashes
		my $hashTemp = $hashFinalSortData{$rank}{$variant};
		my $hashColumn_ref = \%dicoColumnNbr;



#FILL tab 'ALL';
$htmlALL .= "<tr class=\"ALL".$hashFinalSortData{$rank}{$variant}{'worksheet'}."\" >\n";



for( my $fieldNbr = 0 ; $fieldNbr < scalar @{$hashFinalSortData{$rank}{$variant}{'finalArray'}} ; $fieldNbr++){

#foreach my $field ( @{$hashFinalSortData{$rank}{$variant}{'finalArray'}} ){
	#print HTML "\t<td style=\"word-wrap: break-word\";class=\"tooltip\"; title=".$field."  >";
	#print HTML "\t<td >";


	switch ($fieldNbr){
		case ( 0 ) { $htmlALL .= "\t<td ><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentMPAscore'}."</span></div>"   }
		case ( 2 ) { $htmlALL .= "\t<td ><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentPhenolyzer'}."</span></div> "   }
		case ( 3 ) { $htmlALL .= "\t<td style=\"background-color:".$hashFinalSortData{$rank}{$variant}{'colorpLI'}."\"><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentpLI'}."</span></div>    "   }
		case ( 4 ) { $htmlALL .= "\t<td ><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentPhenotype'}."</span></div> "   }
		case ( 5 ) { $htmlALL .= "\t<td ><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'}."</span></div> "   }
		case ( 6 ) { $htmlALL .= "\t<td ><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentGnomADexome'}."</span></div> "   }
		case ( 7 ) { $htmlALL .= "\t<td ><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentClinvar'}."</span></div> "   }
		case ( 8 ) { $htmlALL .= "\t<td ><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentInterVar'}."</span></div> "   }
		case ( 10 ) { $htmlALL .= "\t<td ><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentFunc'}."</span></div> "   }
		case ( 11 ) { $htmlALL .= "\t<td ><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."<span class=\"tooltiptext tooltip-bottom\">".$hashFinalSortData{$rank}{$variant}{'commentGenotype'}."</span></div>"   }
		else{	
			if ( defined $hashColor{$fieldNbr}){
				$htmlALL .= "\t<td style=\"background-color:".$hashColor{$fieldNbr}."\"><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."</div>";    
			}else{
				if (defined $hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]){
					$htmlALL .= "\t<td><div class=\"tooltip\">".$hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]."</div>" ;   
				}else{
					$htmlALL .= "\t<td><div class=\"tooltip\">.</div>" ;   
				}
			}
		}
	}



	#if (defined $hashFinalSortData{$rank}{$variant}{'finalArray'}[$fieldNbr]){
	#}else {
#		$htmlALL .= "."
#	}
	$htmlALL .= "\t</td>\n";

}

#print HTML "</tr>\n";
$htmlALL .= "</tr>\n";





##############################################################
###########################      ALL     #####################

		writeThisSheet ($worksheet,
						$worksheetLine,
						$format_pLI,
						$case,
						$hashTemp,
						$hashColumn_ref	);
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
################# OMIM DOMINANT #############

		if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /OMIMDOM/){

			writeThisSheet ($worksheetOMIMDOM,
							$worksheetLineOMIMDOM,
							$format_pLI,
							$case,
							$hashTemp,
							$hashColumn_ref
						);
			
			$worksheetLineOMIMDOM ++;
		}



##############################################################
################# OMIM RECESSIVE #############

		if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /OMIMREC/){

			writeThisSheet ($worksheetOMIMREC,
							$worksheetLineOMIMREC,
							$format_pLI,
							$case,
							$hashTemp,
							$hashColumn_ref
						);
			
			$worksheetLineOMIMREC ++;
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
#################### GENOTYPE SELECTION ################
##############################################################
		
			
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

		if(defined $trio){
			
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

		}   # END IF TRIO
	
	}   # END FOREACH VARIANT
}	#END FOREACH RANK






#############add autofilters within the sheets

$worksheet->autofilter('A1:Z'.$worksheetLine); # Add autofilter until the end
$worksheetOMIMDOM->autofilter('A1:Z'.$worksheetLineOMIMDOM); # Add autofilter until the end
$worksheetOMIMREC->autofilter('A1:Z'.$worksheetLineOMIMREC); # Add autofilter until the end

if(defined $trio){

	$worksheetHTZcompo->autofilter('A1:Z'.$worksheetLineHTZcompo); # Add autofilter
	$worksheetSNPdadVsCNVmum->autofilter('A1:Z'.$worksheetLineSNPdadVsCNVmum); # Add autofilter
 

	#check inheritance consistency
	if ( log($caseDadVariant/$dadVariant) / log(10) < 0.1 ){
		$worksheetMETA->write( 1, 0, "Dad status : OK (log10(".$caseDadVariant."/".$dadVariant.") < 0.1), Inherited Heterozygous variants Ratio tends toward 0." );
	}else{
		$worksheetMETA->write( 1, 0, "Dad status : BAD (log10(".$caseDadVariant."/".$dadVariant.") > 0.1), Inherited Heterozygous variants Ratio tends toward 0." );
	}

	if ( log($caseMumVariant/$mumVariant) / log(10) < 0.1 ){
		$worksheetMETA->write( 2, 0, "Mum status : OK (log10(".$caseMumVariant."/".$mumVariant.") < 0.1), Inherited Heterozygous variants Ratio tends toward 0." );
	}else{
		$worksheetMETA->write( 2, 0, "Mum status : BAD (log10(".$caseMumVariant."/".$mumVariant.") > 0.1), Inherited Heterozygous variants Ratio tends toward 0." );
	}


} 


#write arguments
$worksheetMETA->write( 4, 0, "Arguments:" );
$worksheetMETA->write( 5, 0, $achabArg );


#write vcf Header 
$worksheetMETA->write( 7, 0, "VCF Header:" );
$worksheetMETA->write( 8, 0, $vcfHeader );



$worksheetSNPmumVsCNVdad->autofilter('A1:Z'.$worksheetLineSNPmumVsCNVdad); # Add autofilter
$worksheetAR->autofilter('A1:Z'.$worksheetLineAR); # Add autofilter
$worksheetDENOVO->autofilter('A1:Z'.$worksheetLineDENOVO); # Add autofilter



if($candidates ne ""){
			$worksheetCandidats->autofilter('A1:Z'.$worksheetLineCandidats);
}





$workbook->close();
close(VCF);


print HTML $htmlStart;
#print HTML $htmlEndTable;
print HTML $htmlALL;
print HTML $htmlEndTable;
print HTML $htmlEnd;


close(HTMl);

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
		$hashColumn_ref) = @_;

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

			#print STDERR "genotype case      ". $hashColumn{'Genotype-'.$case}."\n";
			#print STDERR "nb samples     ".$hashTemp{'nbSample'}."\n";

			$worksheet->write_row( $worksheetLine, 0, $hashTemp{'finalArray'} );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'MPA_ranking'},		$hashTemp{'commentMPAscore'} ,x_scale => 2, y_scale => 5 );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'gnomAD_genome_ALL'},	$hashTemp{'commentGnomADgenome'} ,x_scale => 3, y_scale => 2  );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'gnomAD_exome_ALL'},	$hashTemp{'commentGnomADexome'} ,x_scale => 3, y_scale => 2  );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'Genotype-'.$case},	$hashTemp{'commentGenotype'} ,x_scale => 2, y_scale => $hashTemp{'nbSample'} );
			$worksheet->write_comment( $worksheetLine,$hashColumn{'Func.'.$refGene},		$hashTemp{'commentFunc'} ,x_scale => 3, y_scale => 3  );
			
			#DEBUG print commentGenotype in CHROM cells
			#$worksheet->write_comment( $worksheetLine,$hashColumn{'#CHROM'},	$hashTemp{'commentGenotype'} ,x_scale => 2, y_scale => $hashTemp{'nbSample'} );
			
			
			if ($hashTemp{'commentPhenotype'} ne ""){

				$worksheet->write_comment( $worksheetLine,$hashColumn{'Phenotypes.'.$refGene}, $hashTemp{'commentPhenotype'} ,x_scale => 7, y_scale => 5  );
			}
			
			if ($hashTemp{'commentPhenolyzer'} ne ""){
				$worksheet->write_comment( $worksheetLine,$hashColumn{'Phenolyzer'}, $hashTemp{'commentPhenolyzer'} ,x_scale => 2 );
			}

			if ($hashTemp{'commentInterVar'} ne ""){
				$worksheet->write_comment( $worksheetLine,$hashColumn{'InterVar_automated'}, $hashTemp{'commentInterVar'} ,x_scale => 7, y_scale => 5  );
			}
			
			if ($hashTemp{'commentClinvar'} ne ""){
				$worksheet->write_comment( $worksheetLine,$hashColumn{'CLNSIG'}, $hashTemp{'commentClinvar'} ,x_scale => 7, y_scale => 5  );
			}


			if ($hashTemp{'commentpLI'} ne "."){

        		$format_pLI = $workbook->add_format(bg_color => $hashTemp{'colorpLI'});


				$worksheet->write( $worksheetLine,$hashColumn{'Gene.'.$refGene}, $hashTemp{'finalArray'}[$hashColumn{'Gene.'.$refGene}]     ,$format_pLI );
				$worksheet->write_comment( $worksheetLine,$hashColumn{'Gene.'.$refGene},$hashTemp{'commentpLI'},x_scale => 5, y_scale => 5  );
			}

			if(defined $hashTemp{'genotypeMozaic'} ){	
				my @genotypeMozaic = split (';', $hashTemp{'genotypeMozaic'} );
				#recycling $format_pLI to color mozaic genotypes
				$format_pLI = $workbook->add_format(bg_color => 'purple');
				

				for( my $sampleMozaic = 0 ; $sampleMozaic < scalar @genotypeMozaic; $sampleMozaic +=2){
				#foreach my $sampleMozaic (@genotypeMozaic){
					$format_pLI = $workbook->add_format(bg_color => $genotypeMozaic[$sampleMozaic+1]);
					$worksheet->write( $worksheetLine,$hashColumn{$genotypeMozaic[$sampleMozaic]}  , $hashTemp{'finalArray'}[$hashColumn{$genotypeMozaic[$sampleMozaic]}] ,$format_pLI );
					#$worksheet->write( $worksheetLine,$hashColumn{$sampleMozaic}  , $hashTemp{'finalArray'}[$hashColumn{$sampleMozaic}] ,$format_pLI );
				}	
			}

}#END OF SUB

exit 0; 
