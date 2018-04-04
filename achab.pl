#!/usr/bin/perl

##### vcfToExcel.pl ####

# Auteur : Thomas Guignard 2018
# USAGE : achab.pl --vcf <vcf_file> --cas <index_sample_name> --pere <father_sample_name> --mere <mother_sample_name> --control <control_sample_name>  --output <txt_file> --caller <freebayes|GATK> --trio <YES|NO> --candidats <file with gene symbol of interest>  --geneSummary <file with format geneSymbol\tsummary>  --pLIFile <file containing gene symbol and pLI from ExAC database> --annotation <annovar|seattleseq> --phenolyzerFile <phenolyzer output file suffixed by predicted_gene_scores>
#
#ExAC pLI datas ftp://ftp.broadinstitute.org/pub/ExAC_release/release1/functional_gene_constraint/

# Description : 
# Create an User friendly Excel file from an annotated VCF file. 

# Version : 
# v1.0.0 20161020 Initial Implementation 
# v6 20171201 Use VCF header to get parameter
# v7 20180119 Adapt to annovar annotation
# v8 20180329 Adapt to MPA and phenolyzer outputs

use strict; 
use warnings;
use Getopt::Long; 
use Pod::Usage;
use List::Util qw(first);
use Excel::Writer::XLSX;
use Data::Dumper;


#parameters
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


#stuff for files
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

my $phenolyzerFile = "";
my $phenolyzer_Line;
my @phenolyzer_List;
my %phenolyzerGene;

#vcf parsing and output
my @line;
#my @unorderedLine;
my @orderedLine;

my @finalSortData;
my $familyGenotype;		

	
my $nbrSample = 0;
my $indexPere = 0;
my $indexMere = 0;
my $indexCas = 0;
my $indexControl = 0;

my $annotation = "";

#$arguments = GetOptions( "vcf=s" => \$incfile, "output=s" => \$outfile, "cas=s" => \$cas, "pere=s" => \$pere, "mere=s" => \$mere, "control=s" => \$control );
#$arguments = GetOptions( "vcf=s" => \$incfile ) or pod2usage(-vcf => "$0: argument required\n") ;

GetOptions( 	"vcf=s"				=> \$incfile,
			 	"output=s"			=> \$outfile,
				"cas=s"				=> \$cas,
				"pere=s"			=> \$pere, 
				"mere=s"			=> \$mere, 
				"control=s"			=> \$control,
				"caller=s"			=> \$caller,
				"trio=s"			=> \$trio,
				"candidats=s"		=> \$candidates,
				"geneSummary=s"		=> \$geneSummary,
				"pLIFile=s"			=> \$pLIFile,
				"annotation=s"		=> \$annotation,
				"phenolyzerFile=s"	=> \$phenolyzerFile,
				"man"				=> \$man,
				"help"				=> \$help);
				

#check mandatory arguments



			 
print  STDERR "Processing vcf file ... \n" ; 


open( VCF , "<$incfile" )or die("Cannot open vcf file $incfile") ;
#open(OUT,"| /bin/gzip -c >$outfile".".gz") ;
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
#$worksheet->autofilter('A1:AN5000'); # Add autofilter

my $worksheetACMG = $workbook->add_worksheet('DS_ACMG');
$worksheetACMG->freeze_panes( 1, 0 );    # Freeze the first row
my $worksheetLineACMG = 0;



#my $worksheetOMIM = $workbook->add_worksheet('OMIM');
#$worksheetOMIM->freeze_panes( 1, 0 );    # Freeze the first row
#my $worksheetLineOMIM = 0;

#my $worksheetClinVarPatho = $workbook->add_worksheet('ClinVar_Patho');  
#$worksheetClinVarPatho->freeze_panes( 1, 0 );    # Freeze the first row
#my $worksheetLineClinVarPatho = 0;


#my $worksheetRAREnonPASS = $workbook->add_worksheet('Rare_nonPASS');
#$worksheetRAREnonPASS->freeze_panes( 1, 0 );    # Freeze the first row
#my $worksheetLineRAREnonPASS = 0;


#$worksheetRAREnonPASS->autofilter('A1:AN1000'); # Add autofilter
#$worksheetOMIM->autofilter('A1:AN1000'); # Add autofilter

#$worksheetRAREnonPASS->autofilter('A1:AN1000'); # Add autofilter
#$worksheetOMIM->autofilter('A1:AN1000'); # Add autofilter
#$worksheetClinVarPatho->autofilter('A1:AN1000'); # Add autofilter






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
	#$worksheetHTZcompo->autofilter('A1:AN1000'); # Add autofilter
  
	$worksheetAR = $workbook->add_worksheet('AR');
	$worksheetLineAR = 0;
	$worksheetAR->freeze_panes( 1, 0 );    # Freeze the first row
	#$worksheetAR->autofilter('A1:AN1000'); # Add autofilter
  
	$worksheetSNPmereVsCNVpere = $workbook->add_worksheet('SNVmereVsCNVpere');
	$worksheetLineSNPmereVsCNVpere = 0;
	$worksheetSNPmereVsCNVpere->freeze_panes( 1, 0 );    # Freeze the first row
	#$worksheetSNPmereVsCNVpere->autofilter('A1:AN1000'); # Add autofilter
  
	$worksheetSNPpereVsCNVmere = $workbook->add_worksheet('SNVpereVsCNVmere');
	$worksheetLineSNPpereVsCNVmere = 0;
	$worksheetSNPpereVsCNVmere->freeze_panes( 1, 0 );    # Freeze the first row
	#$worksheetSNPpereVsCNVmere->autofilter('A1:AN1000'); # Add autofilter
  
	$worksheetDENOVO = $workbook->add_worksheet('DENOVO');
	$worksheetLineDENOVO = 0;
	$worksheetDENOVO->freeze_panes( 1, 0 );    # Freeze the first row
	#$worksheetDENOVO->autofilter('A1:AN1000'); # Add autofilter

	#$worksheetDELHMZ = $workbook->add_worksheet('DEL_HMZ');
	#$worksheetLineDELHMZ = 0;
	#$worksheetDELHMZ->freeze_panes( 1, 0 );    # Freeze the first row
	#$worksheetDELHMZ->autofilter('A1:AN1000'); # Add autofilter
		

#$worksheetLineHTZcompo ++;
#$worksheetLineAR ++; 
#$worksheetLineSNPmereVsCNVpere ++;
#$worksheetLineSNPpereVsCNVmere ++;
#$worksheetLineDENOVO ++;

}

#get data from phenolyzer output (predicted_gene_score)
my $current_gene= "";
my $maxLine=0;
if($phenolyzerFile ne ""){
	open(PHENO , "<$phenolyzerFile") or die("Cannot open candidates file $phenolyzerFile") ;
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
	#$worksheetCandidats->autofilter('A1:AN1000'); # Add autofilter
}


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


if($annotation eq "annovar"){
	# evaluate and check if function need to be kept with nsssi
}




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
				
	}
	close(PLIFILE);
}



 

#TODO check if header contains required INFO
#Parse VCF header to fill the dictionnary of parameters
print STDERR "Parsing VCF header to fill the dictionnary of parameters ... \n";
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
		  #DEBUG      print STDERR "info : ". $1 . "\tdescription: ". $2."\n";
			

			    next;
      
      }else {print STDERR "pattern not found in this line: ".$current_line ."\n";next} 
		
    }else {last}

}
close(VCF);



#DEBUG


#counter for shifting columns according to nbr of sample
my $cmpt = 0;

##create Dico of sorted columns for a userfriendly output
#my %dicoOrderColumns;


#$dicoOrderColumns{1 }{'colName'} = $dicoParam{"MPA_ranking"} ;
#$dicoOrderColumns{2 }{'colName'} = $dicoParam{"Gene.refGene"};
#$dicoOrderColumns{3 }{'colName'} = $dicoParam{"CLNDB-Phenotypes"} ;
#$dicoOrderColumns{4 }{'colName'} = $dicoParam{"CLN-Disease_description"};
#$dicoOrderColumns{5 }{'colName'} = $dicoParam{"MIM Number"};
#$dicoOrderColumns{6 }{'colName'} = $dicoParam{"gnomAD_genome_ALL"};
#$dicoOrderColumns{7 }{'colName'} = $dicoParam{"gnomAD_exome_ALL"} ;
#$dicoOrderColumns{8 }{'colName'} = $dicoParam{"CLINSIG"} ;
#$dicoOrderColumns{9 }{'colName'} = $dicoParam{"SecondHit-CNV"} ;
#$dicoOrderColumns{10}{'colName'} = $dicoParam{"Func.refGene"} ;
#$dicoOrderColumns{11}{'colName'} = $dicoParam{"ExonicFunc.refGene"} ;
#$dicoOrderColumns{12}{'colName'} = $dicoParam{"AAChange.refGene"} ;
#$dicoOrderColumns{13}{'colName'} = $dicoParam{"GeneDetail.refGene"} ;
#$dicoOrderColumns{14}{'colName'} = "Genotype-".$cas ;
#if($pere ne "" && $mere ne "" && $control ne ""){
#	$cmpt = 4;
#	
#	$dicoOrderColumns{15}{'colName'} = "Genotype-".$pere;
#	$dicoOrderColumns{16}{'colName'} = "Genotype-".$mere ;
#	$dicoOrderColumns{17}{'colName'} = "Genotype-".$control ;
#	
#
#
#}else{
#	if( $mere ne "" && $pere ne ""){
#		$cmpt = 3;
#		
#		$dicoOrderColumns{15}{'colName'} = "Genotype-".$pere ;
#		$dicoOrderColumns{16}{'colName'} = "Genotype-".$mere ;
#
#	}else{
#		if( $pere ne ""){
#			$cmpt = 2;
#			$dicoOrderColumns{15}{'colName'} = "Genotype-".$pere ;
#		}
#		if( $mere ne ""){
#			$cmpt= 2;
#			$dicoOrderColumns{15}{'colName'} = "Genotype-".$mere ;
#			
#		}
#		if( $mere eq "" && $pere eq ""){
#			$cmpt= 1;
#			$dicoOrderColumns{15}{'colName'} = "Genotype-".$mere ;
#
#	
#	}
#}


#$dicoOrderColumns{14+$cmpt}{'colName'} = "#CHROM" ;
#$dicoOrderColumns{15+$cmpt}{'colName'} = "POS" ;
#$dicoOrderColumns{16+$cmpt}{'colName'} = "ID" ;
#$dicoOrderColumns{17+$cmpt}{'colName'} = "REF" ;
#$dicoOrderColumns{18+$cmpt}{'colName'} = "ALT" ;
#$dicoOrderColumns{19+$cmpt}{'colName'} = "FILTER" ;
#}
#BUG with CA clinicalAssociation from seattle seq, which is absent from vcf in my test 

#Alternative strategy, define only here the final column position
#definition des numeros de colonnes qui prendront les commentaires
#
#

#dico for sample sorting (index / dad / mum / control)
my %dicoSamples ;


#
my %dicoColumnNbr;
$dicoColumnNbr{'MPA_ranking'}=				0;
$dicoColumnNbr{'Phenolyzer'}=				1;
$dicoColumnNbr{'Gene.refGene'}=				2;
$dicoColumnNbr{'CLNDBN'}=					3;
$dicoColumnNbr{'CLNDSDBID'}=				4;
$dicoColumnNbr{'gnomAD_genome_ALL'}=		5;
$dicoColumnNbr{'gnomAD_exome_ALL'}=			6;
$dicoColumnNbr{'CLINSIG'}=					7;
$dicoColumnNbr{'SecondHit-CNV'}=			8;
$dicoColumnNbr{'Func.refGene'}=				9;

$dicoColumnNbr{'ExonicFunc.refGene'}=		10;
$dicoColumnNbr{'AAChange.refGene'}=			11;
$dicoColumnNbr{'GeneDetail.refGene'}=		12;
$dicoColumnNbr{'Genotype-'.$cas}=			13;


if($pere ne "" && $mere ne "" && $control ne ""){
	$cmpt = 4;
	
	$dicoColumnNbr{'Genotype-'.$pere}=		14;
	$dicoColumnNbr{'Genotype-'.$mere}=		15;
	$dicoColumnNbr{'Genotype-'.$control}=	16;

	$dicoSamples{1}{'columnName'} = 'Genotype-'.$cas ;
	$dicoSamples{2}{'columnName'} = 'Genotype-'.$pere ;
	$dicoSamples{3}{'columnName'} = 'Genotype-'.$mere ;
	$dicoSamples{4}{'columnName'} = 'Genotype-'.$control ;





}else{
	if( $mere ne "" && $pere ne ""){
		$cmpt = 3;
		$dicoColumnNbr{'Genotype-'.$pere}=	14;
		$dicoColumnNbr{'Genotype-'.$mere}=	15;

		$dicoSamples{1}{'columnName'} = 'Genotype-'.$cas ;
		$dicoSamples{2}{'columnName'} = 'Genotype-'.$pere ;
		$dicoSamples{3}{'columnName'} = 'Genotype-'.$mere ;


	}else{
		if( $pere ne ""){
			$cmpt = 2;
			$dicoColumnNbr{'Genotype-'.$pere}=14;
		
			$dicoSamples{1}{'columnName'} = 'Genotype-'.$cas ;
			$dicoSamples{2}{'columnName'} = 'Genotype-'.$pere ;
		
		
		}
		if( $mere ne ""){
			$cmpt= 2;
			$dicoColumnNbr{'Genotype-'.$mere}=14;
		
			$dicoSamples{1}{'columnName'} = 'Genotype-'.$cas ;
			$dicoSamples{2}{'columnName'} = 'Genotype-'.$mere ;


		}
		if( $mere eq "" && $pere eq ""){
			$cmpt= 1;
			$dicoSamples{1}{'columnName'} = 'Genotype-'.$cas ;
	
	}
}

$dicoColumnNbr{'#CHROM'}=	13+$cmpt ;
$dicoColumnNbr{'POS'}=		14+$cmpt ;
$dicoColumnNbr{'ID'}=		15+$cmpt ;
$dicoColumnNbr{'REF'}=		16+$cmpt ;
$dicoColumnNbr{'ALT'}=		17+$cmpt ;
$dicoColumnNbr{'FILTER'}=	18+$cmpt ;


#Define column title order
#my $columnTitles="MPA\tGene\tPhenotypes\tDisease\tMIM\tgnomAD_genome_ALL\tgnomAD_exome_ALL\tCLinvar\tSecondHit_CNV\t"

my @columnTitles;
foreach my $key  (sort { $dicoColumnNbr{$a} <=> $dicoColumnNbr{$b} } keys %dicoColumnNbr)  {
	push @columnTitles,  $key;
	#DEBUG print STDERR $key."\n";
}


#Optionnal: precise in which column is the relative comment
#my $columnNbrForComment_MPA=$dicoColumnNbr{'MPA_ranking'};
#my $columnNbrForComment_Gene=$dicoColumnNbr{'Gene.refGene'};
#my $columnNbrForComment_GNOMADGENOME=$dicoColumnNbr{'gnomAD_genome_ALL'};
#my $columnNbrForComment_GNOMADEXOME=$dicoColumnNbr{'gnomAD_exome_ALL'};
#my $columnNbrForComment_Genotype=$dicoColumnNbr{'gnomAD_exome_ALL'};

#define comment content with dico, initialization step
#my %dicoCommentMPAscore;
#my %dicoCommentGnomadExome;
#my %dicoCommentGnomadGenome;
#my %dicoCommentGenotype;

#final strings for comment
my $commentGenotype;
my $commentMPAscore;
my $commentGnomADExomeScore;
my $commentGnomADGenomeScore;

#define sorted arrays with score for comment
my @CommentMPA_score = ("MPA_impact",
						"MPA_adjusted",
						"MPA_available",
						"MPA_deleterious",
						"MPA_final_score",
						"\n---SPLICE---",
						'dbscSNV_ADA_SCORE',
						'dbscSNV_RF_SCORE',
						'dpsi_zscore',
						"\n---MISSENSE---",
						'SIFT_pred',
						'Polyphen2_HDIV_pred',
						'Polyphen2_HVAR_pred',
						'LRT_pred',
						'MutationTaster_pred',
						'FATHMM_pred',
						'fathmm-MKL_coding_pred',
						'PROVEAN_pred',
						'MetaSVM_pred',
						'MetaLR_pred');

#$dicoCommentMPAscore{'MPA_impact'}='.';
#$dicoCommentMPAscore{'MPA_adjusted'}='.';
#$dicoCommentMPAscore{'MPA_available'}='.';
#$dicoCommentMPAscore{'MPA_deleterious'}='.';
#$dicoCommentMPAscore{'MPA_final_score'}='.';

#	$dicoCommentMPAscore{'SIFT_pred'}='.';
#	#	$dicoCommentMPAscore{'Polyphen2_HDIV_pred'}='.';
#	$dicoCommentMPAscore{'Polyphen2_HVAR_pred'}='.';
#	$dicoCommentMPAscore{'LRT_pred'}='.';
#	$dicoCommentMPAscore{'MutationTaster_pred'}='.';
#	$dicoCommentMPAscore{'FATHMM_pred'}='.';
#	$dicoCommentMPAscore{'PROVEAN_pred'}='.';
#	$dicoCommentMPAscore{'fathmm-MKL_coding_pred'}='.';
#	$dicoCommentMPAscore{'MetaSVM_pred'}='.';
#	$dicoCommentMPAscore{'MetaLR_pred'}='.';
#	$dicoCommentMPAscore{'dbscSNV_ADA_SCORE'}='.';
#	$dicoCommentMPAscore{'dbscSNV_RF_SCORE'}='.';
#	$dicoCommentMPAscore{'dpsi_zscore'}='.';

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



#	$dicoCommentGnomadGenome{'gnomAD_genome_ALL'}='.';
#	$dicoCommentGnomadGenome{'gnomAD_genome_AFR'}='.';
#	$dicoCommentGnomadGenome{'gnomAD_genome_AMR'}='.';
#	$dicoCommentGnomadGenome{'gnomAD_genome_ASJ'}='.';
#	$dicoCommentGnomadGenome{'gnomAD_genome_EAS'}='.';
#	$dicoCommentGnomadGenome{'gnomAD_genome_FIN'}='.';
#	$dicoCommentGnomadGenome{'gnomAD_genome_NFE'}='.';
#	$dicoCommentGnomadGenome{'gnomAD_genome_OTH'}='.';

#	$dicoCommentGnomadExome{'gnomAD_exome_ALL'}='.';
#	$dicoCommentGnomadExome{'gnomAD_exome_AFR'}='.';
#	$dicoCommentGnomadExome{'gnomAD_exome_AMR'}='.';
#	$dicoCommentGnomadExome{'gnomAD_exome_ASJ'}='.';
#	$dicoCommentGnomadExome{'gnomAD_exome_EAS'}='.';
#	$dicoCommentGnomadExome{'gnomAD_exome_FIN'}='.';
#	$dicoCommentGnomadExome{'gnomAD_exome_NFE'}='.';
#	$dicoCommentGnomadExome{'gnomAD_exome_OTH'}='.';



#DEBUG this print is useful to check missing parameter in the VCF
#foreach (sort {$a <=> $b} keys %dicoOrderColumns){
# 	print $_."\t".$dicoOrderColumns{$_}{'colName'}."\n";
#}                  



my %dicoGeneForHTZcompo;
my $previousGene ="";

my $indexUnSort = 0;

#############################################
#Start parsing VCF

open( VCF , "<$incfile" )or die("Cannot open vcf file $incfile") ;

#my @colonnes = ("#CHROM","POS","ID","REF","ALT","QUAL","FILTER");




while( <VCF> ){
  	$current_line = $_;

#############################################
#skip header
	next if ($current_line=~/^##/);

	chomp $current_line;
	@line = split( /\t/, $current_line );	
	
#	my $nsssi = ""; # likely useless

	#DEBUG print STDERR "totototototototo1111111111\t". $dicoColumnNbr{'Gene.refGene'}."\n";
#############################################
#Treatment for First line to create header of the output

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
		}
    
    


#########FILLING COLUMN TITLES FOR SHEETS
			$worksheet->write_row( 0, 0, \@columnTitles );
			#$worksheetOMIM->write_row( 0, 0, \@columnTitles );
	    	$worksheetACMG->write_row( 0, 0, \@columnTitles );
		
        #write comment for pLI
	    	$worksheet->write_comment( 0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3 );
	    	$worksheetACMG->write_comment( 0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3  );


	    	$worksheetLine ++;
			#$worksheetLineOMIM ++;
	    	$worksheetLineACMG ++;





			#FILLING COLUMN TITLES FOR TRIO SHEETS
		if ($trio eq "YES"){
			$worksheetHTZcompo->write_row( 0, 0, \@columnTitles );
			$worksheetAR->write_row( 0, 0, \@columnTitles );
			$worksheetSNPmereVsCNVpere->write_row( 0, 0, \@columnTitles );
			$worksheetSNPpereVsCNVmere->write_row( 0, 0, \@columnTitles );
			$worksheetDENOVO->write_row( 0, 0, \@columnTitles );

			
			#write pLI comment
			$worksheetHTZcompo->write_comment(0,$dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);
			$worksheetAR->write_comment(0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);
			$worksheetSNPmereVsCNVpere->write_comment(0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);
			$worksheetSNPpereVsCNVmere->write_comment(0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);
			$worksheetDENOVO->write_comment( 0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3);



			$worksheetLineHTZcompo ++;
			$worksheetLineHTZcompo ++;
			$worksheetLineAR ++; 
			$worksheetLineSNPmereVsCNVpere ++;
			$worksheetLineSNPpereVsCNVmere ++;
			$worksheetLineDENOVO ++;

		}
			
		#FILLING COLUMN TITLES FOR CANDIDATES SHEET

		if($candidates ne ""){
			$worksheetCandidats->write_row( 0, 0, \@columnTitles );
			$worksheetCandidats->write_comment(0, $dicoColumnNbr{'Gene.refGene'}, $pLI_Comment,  x_scale => 3 );
			$worksheetLineCandidats ++;
		}

#DEBUG print STDERR "totototototototo\t". $dicoColumnNbr{'Gene.refGene'}."\n";

		next;
		
#############################################
##############################
##########start to compute variant lines	

	}else {

		#DEBUG 		print STDERR "totototototototo2222\t". $dicoColumnNbr{'Gene.refGene'}."\n";


		#initialise final printable string
		@finalSortData = ("");
		
		my $alt="";
		my $ref="";
#		my %data;

		#Split line with tab 
#		@line = split( /\t/, $current_line );   
					
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
				$dicoInfo{$infoKeyValue[0]} = $infoKeyValue[1];
#DEBUG
#print $infoKeyValue[1]."\n";
			}
		}

#DEBUG
#print Dumper(\%dicoInfo);


		#select only 1% pop freq 
		#TODO check how to treat "." as frequency, And use pop freq threshold as an input parameter (1%)
#		next if(( $dicoInfo{'gnomAD_exome_ALL'} eq ".") || ($dicoInfo{'gnomAD_genome_ALL'} >= 0.01 && $dicoInfo{'gnomAD_exome_ALL'} >= 0.01));  


		#filling output line, check if INFO exists in the VCF
		foreach my $keys (sort keys %dicoColumnNbr){
			
#			print "keysListe\t#".$keys."#\n";

			if (defined $dicoInfo{$keys}){
				
				$finalSortData[$dicoColumnNbr{$keys}] = $dicoInfo{$keys};
				#DEBUG
				#print "finalSort\t".$finalSortData[$dicoColumnNbr{$keys}]."\n";
				#DEBUG
				#print "dicoInfo\t".$dicoInfo{$keys}."\n";
				#DEBUG
				#print "keys\t".$keys."\n";
			}
		}	

		#Phenolyzer Column
		if($phenolyzerFile ne ""){
			if (defined $phenolyzerGene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]} ){
				$finalSortData[$dicoColumnNbr{'Phenolyzer'}] = $phenolyzerGene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'Raw'};
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
		
		
		
		################ continue to fill the others dicoComment

		


#############	Check FORMAT related to caller 
		#		GT:AD:DP:GQ:PL (haplotype caller)
		#		GT:DP:GQ  => multiallelic line , vcf not splitted => should be done before, STOP RUN?
		#		GT:GOF:GQ:NR:NV:PL (platyplus caller)
		if($line[8] eq "GT:AD:DP:GQ:PL"){
			$caller = "GATK";
		}elsif($line[8] eq "GT:GOF:GQ:NR:NV:PL"){
			$caller = "platypus";
		}elsif($line[8] eq "GT:DP:GQ"){
			print STDERR "Multi-allelic line detection. Please split the vcf file in order to get 1 allele by line\n";
			print STDERR $current_line ."\n";
			exit 1; 
		}else{
			print STDERR "The Caller used for this line is unknown. Treating line is a risky business.\n";
			print STDERR $current_line ."\n";
			exit 1; 
		}
		


#############   Parse Genotypes
		#genotype concatenation for easy hereditary status
		$familyGenotype = "_";		
		$commentGenotype = "CALLER = ".$caller."\tQUALITY = ".$line[5]."\n\n";


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

						}else {
							$DP = $genotype[2];
							$AD = $genotype[1];
						
						}

					}elsif($caller eq "platypus"){
						
						$DP = $genotype[3];
						$AD = ($genotype[3] - $genotype[4]);
						$AD .= ",".$genotype[4];
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
				$commentGenotype .=  $dicoSamples{$finalcol}{'columnName'}."\t-\t".$genotype[0]."\nDP = ".$DP."\tAD = ".$AD."\tAB = ".$AB."\n\n";
				$familyGenotype .= $genotype[0]."_";
		
		} #END of Sample Treatment	

	} #END of IF-ELSE(#CHROM)	





#DEBUG			print Dumper(\%dicoColumnNbr);
##############check hereditary hypothesis or genes to fill sheets

############ TIME TO FILL THE XLSX OUTPUT FILE


			#Concatenate  Gene summary with gene Name
			if(defined $dicoGeneSummary{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}){
				$geneSummaryConcat = $finalSortData[$dicoColumnNbr{'Gene.refGene'}]." ### ".$dicoGeneSummary{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]};
			}else{
		        	$geneSummaryConcat = $finalSortData[$dicoColumnNbr{'Gene.refGene'}];
			}


			#Add color code format related to pLI
			if(defined $dicopLI{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{"comment"}){
				$pLI_values = $dicopLI{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{"comment"};
        		$format_pLI = $workbook->add_format(bg_color => $dicopLI{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{"color_format"});

			}else{
				$pLI_values = "";
              			#$format_pLI -> set_bg_color('#FFFFFF');
				$format_pLI = $workbook->add_format(bg_color => '#FFFFFF');

			} 




			#write incidentalome-ACMG sheet
			foreach my $ACMGgene (@ACMGlist) {
				if($finalSortData[$dicoColumnNbr{'Gene.refGene'}] eq $ACMGgene){
					$worksheetACMG->write_row( $worksheetLineACMG, 0, \@finalSortData );
					$worksheetACMG->write( $worksheetLineACMG,$dicoColumnNbr{'Gene.refGene'},$geneSummaryConcat,$format_pLI );
					if($pLI_values ne ""){
						$worksheetACMG->write_comment( $worksheetLineACMG,$dicoColumnNbr{'Gene.refGene'},$pLI_values,x_scale => 2 );
					}	
					$worksheetLineACMG ++;
				}
			}
			
			#write candidates genes sheet
			if($candidates ne ""){
				foreach my $candGene (@candidatesList){
					if($finalSortData[$dicoColumnNbr{'Gene.refGene'}] eq $candGene){
						$worksheetCandidats->write_row( $worksheetLineCandidats, 0, \@finalSortData );
						$worksheetCandidats->write( $worksheetLineCandidats,$dicoColumnNbr{'Gene.refGene'},$geneSummaryConcat,$format_pLI );
						if($pLI_values ne ""){
							$worksheetCandidats->write_comment( $worksheetLineCandidats,$dicoColumnNbr{'Gene.refGene'},$pLI_values ,x_scale => 2);
						}	
						$worksheetLineCandidats ++;
					}
				}
			}


			#additionnal analysis in TRIO context
			if ($trio eq "YES"){

				
#################maybe consider also 1/0 genotype or correct it before
		
				#DEBUG print $familyGenotype."\n";

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

						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@emptyArray );
						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@finalSortData );
						$worksheetHTZcompo->write( $worksheetLineHTZcompo, $dicoColumnNbr{'Gene.refGene'},$geneSummaryConcat,$format_pLI );
						if($pLI_values ne ""){
							$worksheetHTZcompo->write_comment( $worksheetLineHTZcompo,$dicoColumnNbr{'Gene.refGene'},$pLI_values,x_scale => 2);
						}
						$worksheetLineHTZcompo ++;

					}else{

#						if(defined $dicoGeneForHTZcompo{$previousGene} && $dicoGeneForHTZcompo{$previousGene}{'ok'}==0 ){
						if(($previousGene ne $finalSortData[$dicoColumnNbr{'Gene.refGene'}]) && $dicoGeneForHTZcompo{$previousGene}{'ok'}==0 ){
							$worksheetLineHTZcompo -= $dicoGeneForHTZcompo{$previousGene}{'cnt'};
							$worksheetAR->write_comment( $worksheetLineAR, $dicoColumnNbr{'Gene.refGene'},$pLI_values ,x_scale => 2);
						}		
						$worksheetLineAR ++; 
				}


				#Find SNVvsCNV
				if($familyGenotype eq "_1/1_0/0_0/1_"   || $familyGenotype eq "_1/1_0/0_1/0_"){
					$worksheetSNPmereVsCNVpere->write_row($worksheetLineSNPmereVsCNVpere , 0, \@finalSortData );
					$worksheetSNPmereVsCNVpere->write( $worksheetLineSNPmereVsCNVpere, $dicoColumnNbr{'Gene.refGene'},$geneSummaryConcat,$format_pLI );
					if($pLI_values ne ""){
						$worksheetSNPmereVsCNVpere->write_comment( $worksheetLineSNPmereVsCNVpere,$dicoColumnNbr{'Gene.refGene'} ,$pLI_values,x_scale => 2 );
					}
					$worksheetLineSNPmereVsCNVpere ++;
				}

				if($familyGenotype eq "_1/1_0/1_0/0_" || $familyGenotype eq "_1/1_1/0_0/0_" ){
					$worksheetSNPpereVsCNVmere->write_row($worksheetLineSNPpereVsCNVmere , 0, \@finalSortData  );
					$worksheetSNPpereVsCNVmere->write( $worksheetLineSNPpereVsCNVmere, $dicoColumnNbr{'Gene.refGene'},$geneSummaryConcat,$format_pLI );
					if($pLI_values ne ""){
						$worksheetSNPpereVsCNVmere->write_comment( $worksheetLineSNPpereVsCNVmere, $dicoColumnNbr{'Gene.refGene'},$pLI_values ,x_scale => 2);
					}
					$worksheetLineSNPpereVsCNVmere ++;
				}

				#Find de novo
				if($familyGenotype eq "_1/1_0/0_0/0_" || $familyGenotype eq "_0/1_0/0_0/0_"|| $familyGenotype eq "_1/0_0/0_0/0_" ){
					$worksheetDENOVO->write_row( $worksheetLineDENOVO, 0, \@finalSortData );
					$worksheetDENOVO->write( $worksheetLineDENOVO, $dicoColumnNbr{'Gene.refGene'},$geneSummaryConcat,$format_pLI );
					if($pLI_values ne ""){
						$worksheetDENOVO->write_comment( $worksheetLineDENOVO, $dicoColumnNbr{'Gene.refGene'},$pLI_values,x_scale => 2 );
					}
					$worksheetLineDENOVO ++;
				}

				


			}
			#DEBUG print "toto\n";
			#DEBUG print Dumper(\@finalSortData);
			




      #write sheet "ALL"
		#write complete line
			$worksheet->write_row( $worksheetLine, 0, \@finalSortData );
		#write genesummary
		#	
			if($geneSummary ne ""){
				$worksheet->write( $worksheetLine, $dicoColumnNbr{'Gene.refGene'},$geneSummaryConcat,$format_pLI );
			}

			#add comment for MPA column
			$worksheet->write_comment( $worksheetLine, $dicoColumnNbr{'MPA_ranking'}, $commentMPAscore ,x_scale => 3);
			#add comment for genotype
			$worksheet->write_comment( $worksheetLine, $dicoColumnNbr{'Genotype-'.$cas}, $commentGenotype ,x_scale => 3);
			#add comment for gnomAD_genome
			$worksheet->write_comment( $worksheetLine, $dicoColumnNbr{'gnomAD_genome_ALL'}, $commentGnomADGenomeScore ,x_scale => 3);
			#add comment for gnomAD_exome
			$worksheet->write_comment( $worksheetLine, $dicoColumnNbr{'gnomAD_exome_ALL'}, $commentGnomADExomeScore ,x_scale => 3);
			#add comment for phenolyzer
			if(defined $finalSortData[$dicoColumnNbr{'Phenolyzer'}]){
				$worksheet->write_comment( $worksheetLine, $dicoColumnNbr{'Phenolyzer'},$phenolyzerGene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'comment'} ,x_scale => 2);
			}
			#add pLI comment 
			if($pLI_values ne ""){
				$worksheet->write_comment( $worksheetLine, $dicoColumnNbr{'Gene.refGene'},$pLI_values ,x_scale => 2);
			}
			$worksheetLine ++;

		}

	}
}





#add autofilters within the sheets
$worksheet->autofilter('A1:Z'.$worksheetLine); # Add autofilter until the end

if($trio eq "YES"){

	$worksheetHTZcompo->autofilter('A1:Z'.$worksheetLineHTZcompo); # Add autofilter
  
	$worksheetAR->autofilter('A1:Z'.$worksheetLineAR); # Add autofilter
  
	$worksheetSNPmereVsCNVpere->autofilter('A1:Z'.$worksheetLineSNPmereVsCNVpere); # Add autofilter
  
	$worksheetSNPpereVsCNVmere->autofilter('A1:Z'.$worksheetLineSNPpereVsCNVmere); # Add autofilter
  
	$worksheetDENOVO->autofilter('A1:Z'.$worksheetLineDENOVO); # Add autofilter


}

if($candidates ne ""){
			$worksheetCandidats->autofilter('A1:Z'.$worksheetLineCandidats);
}


close(VCF);
print STDERR "Done!";

exit 0;