#!/usr/bin/perl

##### achab.pl ####

# Auteur : Thomas Guignard 2018
# USAGE : achab.pl --vcf <vcf_file> --cas <index_sample_name> --pere <father_sample_name> --mere <mother_sample_name> --control <control_sample_name>  --caller <freebayes|GATK> --trio <YES|NO> --candidats <file with gene symbol of interest>  --phenolyzerFile <phenolyzer output file suffixed by predicted_gene_scores>   --popFreqMax <allelic frequency threshold from 0 to 1 default=0.02>  --customInfo  <info name (will be added in a new column)>
#

# Description : 
# Create an User friendly Excel file from an MPA annotated VCF file. 

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
use Switch;

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
my $popFreqMax = "";

#stuff for files
my $candidates = "";
my %candidateGene;
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

my $customInfo = "";

#vcf parsing and output
my @line;
my $variantID; 

my @finalSortData;
my $familyGenotype;		
my %hashFinalSortData;
	
my $nbrSample = 0;
my $indexPere = 0;
my $indexMere = 0;
my $indexCas = 0;
my $indexControl = 0;

my $annotation = "";

#$arguments = GetOptions( "vcf=s" => \$incfile, "output=s" => \$outfile, "cas=s" => \$cas, "pere=s" => \$pere, "mere=s" => \$mere, "control=s" => \$control );
#$arguments = GetOptions( "vcf=s" => \$incfile ) or pod2usage(-vcf => "$0: argument required\n") ;

GetOptions( 	"vcf=s"				=> \$incfile,
				"cas=s"				=> \$cas,
				"pere=s"			=> \$pere, 
				"mere=s"			=> \$mere, 
				"control=s"			=> \$control,
				"caller=s"			=> \$caller,
				"trio=s"			=> \$trio,
				"candidats=s"		=> \$candidates,
				"geneSummary=s"		=> \$geneSummary,
				"pLIFile=s"			=> \$pLIFile,
				"phenolyzerFile=s"	=> \$phenolyzerFile,
				"popFreqMax=s"		=> \$popFreqMax, 
				"customInfo=s"		=> \$customInfo, 
				"man"				=> \$man,
				"help"				=> \$help);
				

#check mandatory arguments

			#define popFreqMax
if( $popFreqMax eq ""){
	$popFreqMax = 0.01;
	
}

			 
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
		$candidateGene{$candidates_line} = 1;

				
	}
	close(CANDIDATS);
	$worksheetCandidats = $workbook->add_worksheet('Candidats');
	$worksheetCandidats->freeze_panes( 1, 0 );    # Freeze the first row
	#$worksheetCandidats->autofilter('A1:AN1000'); # Add autofilter
}


#create dico of functions for SNV and indel
#my %dicoFonction;
#$dicoFonction{"coding-unknown"} = 0 ;
#$dicoFonction{"coding-unknown-near-splice"} = 0 ;
#$dicoFonction{"missense"} = 0 ;
#$dicoFonction{"missense-near-splice"} = 0 ;
#$dicoFonction{"splice-acceptor"} = 0 ;
#$dicoFonction{"splice-donor"} = 0 ;
#$dicoFonction{"stop-gained"} = 0 ;
#$dicoFonction{"stop-gained-near-splice"} = 0 ;
#$dicoFonction{"stop-lost"} = 0 ;
#$dicoFonction{"stop-lost-near-splice"} = 0 ;
#$dicoFonction{"coding"} = 0 ;
#$dicoFonction{"coding-near-splice"} = 0 ;
#$dicoFonction{"codingComplex"} = 0 ;
#$dicoFonction{"codingComplex-near-splice"} = 0 ;
#$dicoFonction{"frameshift"} = 0 ;
#$dicoFonction{"synonymous-near-splice"} = 0 ;
#$dicoFonction{"intron-near-splice"} = 0 ;
#$dicoFonction{"frameshift-near-splice"} = 0 ;





#Hash of ACMG incidentalome genes
#my @ACMGlist = ("ACTA2","ACTC1","APC","APOB","ATP7B","BMPR1A","BRCA1","BRCA2","CACNA1S","COL3A1","DSC2","DSG2","DSP","FBN1","GLA","KCNH2","KCNQ1","LDLR","LMNA","MEN1","MLH1","MSH2","MSH6","MUTYH","MYBPC3","MYH11","MYH7","MYL2","MYL3","NF2","OTC","PCSK9","PKP2","PMS2","PRKAG2","PTEN","RB1","RET","RYR1","RYR2","SCN5A","SDHAF2","SDHB","SDHC","SDHD","SMAD3","SMAD4","STK11","TGFBR1","TGFBR2","TMEM43","TNNI3","TNNT2","TP53","TPM1","TSC1","TSC2","VHL","WT1");
my %ACMGgene = ("ACTA2" =>1,"ACTC1" =>1,"APC" =>1,"APOB" =>1,"ATP7B" =>1,"BMPR1A" =>1,"BRCA1" =>1,"BRCA2" =>1,"CACNA1S" =>1,"COL3A1" =>1,"DSC2" =>1,"DSG2" =>1,"DSP" =>1,"FBN1" =>1,"GLA" =>1,"KCNH2" =>1,"KCNQ1" =>1,"LDLR" =>1,"LMNA" =>1,"MEN1" =>1,"MLH1" =>1,"MSH2" =>1,"MSH6" =>1,"MUTYH" =>1,"MYBPC3" =>1,"MYH11" =>1,"MYH7" =>1,"MYL2" =>1,"MYL3" =>1,"NF2" =>1,"OTC" =>1,"PCSK9" =>1,"PKP2" =>1,"PMS2" =>1,"PRKAG2" =>1,"PTEN" =>1,"RB1" =>1,"RET" =>1,"RYR1" =>1,"RYR2" =>1,"SCN5A" =>1,"SDHAF2" =>1,"SDHB" =>1,"SDHC" =>1,"SDHD" =>1,"SMAD3" =>1,"SMAD4" =>1,"STK11" =>1,"TGFBR1" =>1,"TGFBR2" =>1,"TMEM43" =>1,"TNNI3" =>1,"TNNT2" =>1,"TP53" =>1,"TPM1" =>1,"TSC1" =>1,"TSC2" =>1,"VHL" =>1,"WT1"=>1);



                  
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





#counter for shifting columns according to nbr of sample
my $cmpt = 0;


#dico for sample sorting (index / dad / mum / control)
my %dicoSamples ;


#
my %dicoColumnNbr;
$dicoColumnNbr{'MPA_ranking'}=				0;
$dicoColumnNbr{'Phenolyzer'}=				1;
$dicoColumnNbr{'Gene.refGene'}=				2;
$dicoColumnNbr{'CLNDBN'}=					3;  #OMIM - phenotypes
$dicoColumnNbr{'CLNDSDBID'}=				4;	#phenotypes 
												#tissue specificity
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

#Add custom Info in a new column
if($customInfo ne ""){
	$dicoColumnNbr{$customInfo}=	19+$cmpt ;
}



#Define column title order
#my $columnTitles="MPA\tGene\tPhenotypes\tDisease\tMIM\tgnomAD_genome_ALL\tgnomAD_exome_ALL\tCLinvar\tSecondHit_CNV\t"

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







my %dicoGeneForHTZcompo;
my $previousGene ="";

#my $indexUnSort = 0;

#############################################
#Start parsing VCF

open( VCF , "<$incfile" )or die("Cannot open vcf file $incfile") ;


while( <VCF> ){
  	$current_line = $_;

#############################################
#skip header
	next if ($current_line=~/^##/);

	chomp $current_line;
	@line = split( /\t/, $current_line );	
	
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


		next;
		
#############################################
##############################
##########start to compute variant lines	

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
				$dicoInfo{$infoKeyValue[0]} = $infoKeyValue[1];
				#DEBUG
				#print $infoKeyValue[1]."\n";
			}
		}

#DEBUG
#print Dumper(\%dicoInfo);


		#select only x% pop freq 
		#Use pop freq threshold as an input parameter (default = 2%)
		next if(( $dicoInfo{'gnomAD_genome_ALL'} ne ".") && ($dicoInfo{'gnomAD_genome_ALL'} >= $popFreqMax));  


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
			
			#refine MPA_rank for rank 7 missense 
			if($keys eq "MPA_final_score" &&  $finalSortData[$dicoColumnNbr{'MPA_ranking'}] == 7 ){
				$finalSortData[$dicoColumnNbr{'MPA_ranking'}] += (10-$dicoInfo{$keys})/100;
				#print $finalSortData[$dicoColumnNbr{'MPA_ranking'}]."\n";
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


		
		#	print Dumper(\@finalSortData);

#########FILL HASH STRUCTURE FOR FINAL SORT AND OUTPUT, according to rank
		
		#concatenate chrom_POS_REF_ALT to get variant ID
		$variantID = $line[0]."_".$line[1]."_".$line[3]."_".$line[4];
		
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'finalArray'} = [@finalSortData] ; 
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentGnomADexome'} = $commentGnomADExomeScore  ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentGnomADgenome'} = $commentGnomADGenomeScore ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentGenotype'} = $commentGenotype ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentMPAscore'} = $commentMPAscore  ;

			#initialize worksheet
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} = "";


			#ACMG
			if(defined $ACMGgene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]} ){
				$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} = "ACMG";
			}
		
			#CANDIDATES
			if($candidates ne ""){
				if(defined $candidateGene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]} ){
					$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'worksheet'} .= "#CANDIDATES";
				
				}
				
			}
			
			#PHENOLYZER COMMENT
			if(defined  $phenolyzerGene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}){
				$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentPhenolyzer'} = $phenolyzerGene{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{'comment'}  ;

			}





		#additionnal analysis in TRIO context
			if ($trio eq "YES"){
			
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




		#create pLI comment and format
		if(defined $dicoInfo{'pLi.refGene'} && $dicoInfo{'pLi.refGene'} ne "." ){

			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} =  "pLI = ".$dicoInfo{'pLi.refGene'}."\npRec = ".$dicoInfo{'pRec.refGene'}."\npNull = ".$dicoInfo{'pNull.refGene'} ;
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'colorpLI'} =  sprintf('#%2.2X%2.2X%2.2X',($dicoInfo{'pLi.refGene'}*255 + $dicoInfo{'pRec.refGene'}*255),($dicoInfo{'pRec.refGene'}*255 + $dicoInfo{'pNull.refGene'} * 255),0) ;

        	$format_pLI = $workbook->add_format(bg_color => $dicopLI{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{"color_format"});
		
		}else{	
			
			$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} =  "." ;
			
			$format_pLI = $workbook->add_format(bg_color => '#FFFFFF');
		}



	} #END of IF-ELSE(#CHROM)	





#DEBUG			print Dumper(\%dicoColumnNbr);
##############check hereditary hypothesis or genes to fill sheets

############ TIME TO FILL THE XLSX OUTPUT FILE


			#Concatenate  Gene summary with gene Name
#			if(defined $dicoGeneSummary{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}){
#				$geneSummaryConcat = $finalSortData[$dicoColumnNbr{'Gene.refGene'}]." ### ".$dicoGeneSummary{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]};
#			}else{
#		        	$geneSummaryConcat = $finalSortData[$dicoColumnNbr{'Gene.refGene'}];
#			}


#			#Add color code format related to pLI
#			if(defined $dicopLI{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{"comment"}){
#				$pLI_values = $dicopLI{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{"comment"};
			#       		$format_pLI = $workbook->add_format(bg_color => $dicopLI{$finalSortData[$dicoColumnNbr{'Gene.refGene'}]}{"color_format"});

#			}else{
#				$pLI_values = "";
#			
#				$format_pLI = $workbook->add_format(bg_color => '#FFFFFF');

#			} 



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
						if($hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} ne "."){
								$worksheetHTZcompo->write( $worksheetLineHTZcompo, $dicoColumnNbr{'Gene.refGene'},$finalSortData[$dicoColumnNbr{'Gene.refGene'}],$format_pLI );
								$worksheetHTZcompo->write_comment( $worksheetLineHTZcompo,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'},x_scale => 2);
							
						}
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

						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@emptyArray );
						$worksheetHTZcompo->write_row( $worksheetLineHTZcompo, 0, \@finalSortData );
																																									
						if($hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'} ne "."){
								$worksheetHTZcompo->write( $worksheetLineHTZcompo, $dicoColumnNbr{'Gene.refGene'},$finalSortData[$dicoColumnNbr{'Gene.refGene'}],$format_pLI  );
								$worksheetHTZcompo->write_comment( $worksheetLineHTZcompo,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$finalSortData[$dicoColumnNbr{'MPA_ranking'}]}{$variantID}{'commentpLI'},x_scale => 2);
						}	
						
						
																																									$worksheetLineHTZcompo ++;
																																									$previousGene = $finalSortData[$dicoColumnNbr{'Gene.refGene'}];
						
						}
					}

				}#END IF HTZ COMPO
			
			}# END IF TRIO

}#END WHILE VCF


#try to output sort by rank

foreach my $rank (sort {$a <=> $b} keys %hashFinalSortData){
	#print $rank."\n";
	
	foreach my $variant ( keys %{$hashFinalSortData{$rank}}){

		#print $variant."\n";

		###############  ALL #####################
			$format_pLI = $workbook->add_format(bg_color => '#FFFFFF');
 
			
			$worksheet->write_row( $worksheetLine, 0, $hashFinalSortData{$rank}{$variant}{'finalArray'} );
			$worksheet->write_comment( $worksheetLine,$dicoColumnNbr{'MPA_ranking'}, $hashFinalSortData{$rank}{$variant}{'commentMPAscore'} ,x_scale => 2 );
			$worksheet->write_comment( $worksheetLine,$dicoColumnNbr{'gnomAD_genome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'} ,x_scale => 2 );
			$worksheet->write_comment( $worksheetLine,$dicoColumnNbr{'gnomAD_exome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADexome'} ,x_scale => 2 );
			$worksheet->write_comment( $worksheetLine,$dicoColumnNbr{'Genotype-'.$cas}, $hashFinalSortData{$rank}{$variant}{'commentGenotype'} ,x_scale => 2 );
			$worksheet->write_comment( $worksheetLine,$dicoColumnNbr{'Phenolyzer'}, $hashFinalSortData{$rank}{$variant}{'commentPhenolyzer'} ,x_scale => 2 );


			if ($hashFinalSortData{$rank}{$variant}{'commentpLI'} ne "."){

        		$format_pLI = $workbook->add_format(bg_color => $hashFinalSortData{$rank}{$variant}{'colorpLI'});


				$worksheet->write( $worksheetLine,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'finalArray'}[$dicoColumnNbr{'Gene.refGene'}]     ,$format_pLI );
				$worksheet->write_comment( $worksheetLine,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'commentpLI'},x_scale => 2 );
			}	

			$worksheetLine ++;
	






		######### ACMG DS #############

		if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /ACMG/){

			$worksheetACMG->write_row( $worksheetLineACMG, 0, $hashFinalSortData{$rank}{$variant}{'finalArray'} );
			$worksheetACMG->write_comment( $worksheetLineACMG,$dicoColumnNbr{'MPA_ranking'}, $hashFinalSortData{$rank}{$variant}{'commentMPAscore'} ,x_scale => 2 );
			$worksheetACMG->write_comment( $worksheetLineACMG,$dicoColumnNbr{'gnomAD_genome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'} ,x_scale => 2 );
			$worksheetACMG->write_comment( $worksheetLineACMG,$dicoColumnNbr{'gnomAD_exome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADexome'} ,x_scale => 2 );
			$worksheetACMG->write_comment( $worksheetLineACMG,$dicoColumnNbr{'Genotype-'.$cas}, $hashFinalSortData{$rank}{$variant}{'commentGenotype'} ,x_scale => 2 );
			$worksheetACMG->write_comment( $worksheetLineACMG,$dicoColumnNbr{'Phenolyzer'}, $hashFinalSortData{$rank}{$variant}{'commentPhenolyzer'} ,x_scale => 2 );


			if ($hashFinalSortData{$rank}{$variant}{'commentpLI'} ne "."){
				$worksheetACMG->write( $worksheetLineACMG,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'finalArray'}[$dicoColumnNbr{'Gene.refGene'}]     ,$format_pLI );
				$worksheetACMG->write_comment( $worksheetLineACMG,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'commentpLI'},x_scale => 2 );
			}	

			$worksheetLineACMG ++;
		}
	


		######### CANDIDATES #############

		if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /CANDIDATE/){

			$worksheetCandidats->write_row( $worksheetLineCandidats, 0, $hashFinalSortData{$rank}{$variant}{'finalArray'} );
			$worksheetCandidats->write_comment( $worksheetLineCandidats,$dicoColumnNbr{'MPA_ranking'}, $hashFinalSortData{$rank}{$variant}{'commentMPAscore'} ,x_scale => 2 );
			$worksheetCandidats->write_comment( $worksheetLineCandidats,$dicoColumnNbr{'gnomAD_genome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'} ,x_scale => 2 );
			$worksheetCandidats->write_comment( $worksheetLineCandidats,$dicoColumnNbr{'gnomAD_exome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADexome'} ,x_scale => 2 );
			$worksheetCandidats->write_comment( $worksheetLineCandidats,$dicoColumnNbr{'Genotype-'.$cas}, $hashFinalSortData{$rank}{$variant}{'commentGenotype'} ,x_scale => 2 );
			$worksheetCandidats->write_comment( $worksheetLineCandidats,$dicoColumnNbr{'Phenolyzer'}, $hashFinalSortData{$rank}{$variant}{'commentPhenolyzer'} ,x_scale => 2 );


			if ($hashFinalSortData{$rank}{$variant}{'commentpLI'} ne "."){
				$worksheetCandidats->write( $worksheetLineCandidats,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'finalArray'}[$dicoColumnNbr{'Gene.refGene'}]     ,$format_pLI );
				$worksheetCandidats->write_comment( $worksheetLineCandidats,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'commentpLI'},x_scale => 2 );
			}	

			$worksheetLineCandidats ++;
		}
	

		if($trio eq "YES"){
			
			#DENOVO
			if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /DENOVO/){

				$worksheetDENOVO->write_row( $worksheetLineDENOVO, 0, $hashFinalSortData{$rank}{$variant}{'finalArray'} );
				$worksheetDENOVO->write_comment( $worksheetLineDENOVO,$dicoColumnNbr{'MPA_ranking'}, $hashFinalSortData{$rank}{$variant}{'commentMPAscore'} ,x_scale => 2 );
				$worksheetDENOVO->write_comment( $worksheetLineDENOVO,$dicoColumnNbr{'gnomAD_genome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'} ,x_scale => 2 );
				$worksheetDENOVO->write_comment( $worksheetLineDENOVO,$dicoColumnNbr{'gnomAD_exome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADexome'} ,x_scale => 2 );
				$worksheetDENOVO->write_comment( $worksheetLineDENOVO,$dicoColumnNbr{'Genotype-'.$cas}, $hashFinalSortData{$rank}{$variant}{'commentGenotype'} ,x_scale => 2 );
				$worksheetDENOVO->write_comment( $worksheetLineDENOVO,$dicoColumnNbr{'Phenolyzer'}, $hashFinalSortData{$rank}{$variant}{'commentPhenolyzer'} ,x_scale => 2 );


				if ($hashFinalSortData{$rank}{$variant}{'commentpLI'} ne "."){
					$worksheetDENOVO->write( $worksheetLineDENOVO,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'finalArray'}[$dicoColumnNbr{'Gene.refGene'}]     ,$format_pLI );
					$worksheetDENOVO->write_comment( $worksheetLineDENOVO,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'commentpLI'},x_scale => 2 );
				}	

				$worksheetLineDENOVO ++;

				next;

			}	


			#AUTOSOMIC RECESSIVE HOMOZYGOUS
			if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /AUTOREC/){

				$worksheetAR->write_row( $worksheetLineAR, 0, $hashFinalSortData{$rank}{$variant}{'finalArray'} );
				$worksheetAR->write_comment( $worksheetLineAR,$dicoColumnNbr{'MPA_ranking'}, $hashFinalSortData{$rank}{$variant}{'commentMPAscore'} ,x_scale => 2 );
				$worksheetAR->write_comment( $worksheetLineAR,$dicoColumnNbr{'gnomAD_genome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'} ,x_scale => 2 );
				$worksheetAR->write_comment( $worksheetLineAR,$dicoColumnNbr{'gnomAD_exome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADexome'} ,x_scale => 2 );
				$worksheetAR->write_comment( $worksheetLineAR,$dicoColumnNbr{'Genotype-'.$cas}, $hashFinalSortData{$rank}{$variant}{'commentGenotype'} ,x_scale => 2 );
				$worksheetAR->write_comment( $worksheetLineAR,$dicoColumnNbr{'Phenolyzer'}, $hashFinalSortData{$rank}{$variant}{'commentPhenolyzer'} ,x_scale => 2 );


				if ($hashFinalSortData{$rank}{$variant}{'commentpLI'} ne "."){
					$worksheetAR->write( $worksheetLineAR,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'finalArray'}[$dicoColumnNbr{'Gene.refGene'}]     ,$format_pLI );
					$worksheetAR->write_comment( $worksheetLineAR,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'commentpLI'},x_scale => 2 );
				}	
				
				$worksheetLineAR ++;

				next;

			}

			#SNPvsCNV
			if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /SNPpCNVm/) {

				$worksheetSNPpereVsCNVmere->write_row( $worksheetLineSNPpereVsCNVmere, 0, $hashFinalSortData{$rank}{$variant}{'finalArray'} );
				$worksheetSNPpereVsCNVmere->write_comment( $worksheetLineSNPpereVsCNVmere,$dicoColumnNbr{'MPA_ranking'}, $hashFinalSortData{$rank}{$variant}{'commentMPAscore'} ,x_scale => 2 );
				$worksheetSNPpereVsCNVmere->write_comment( $worksheetLineSNPpereVsCNVmere,$dicoColumnNbr{'gnomAD_genome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'} ,x_scale => 2 );
				$worksheetSNPpereVsCNVmere->write_comment( $worksheetLineSNPpereVsCNVmere,$dicoColumnNbr{'gnomAD_exome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADexome'} ,x_scale => 2 );
				$worksheetSNPpereVsCNVmere->write_comment( $worksheetLineSNPpereVsCNVmere,$dicoColumnNbr{'Genotype-'.$cas}, $hashFinalSortData{$rank}{$variant}{'commentGenotype'} ,x_scale => 2 );
				$worksheetSNPpereVsCNVmere->write_comment( $worksheetLineSNPpereVsCNVmere,$dicoColumnNbr{'Phenolyzer'}, $hashFinalSortData{$rank}{$variant}{'commentPhenolyzer'} ,x_scale => 2 );


				if ($hashFinalSortData{$rank}{$variant}{'commentpLI'} ne "."){
					$worksheetSNPpereVsCNVmere->write( $worksheetLineSNPpereVsCNVmere,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'finalArray'}[$dicoColumnNbr{'Gene.refGene'}]     ,$format_pLI );
					$worksheetSNPpereVsCNVmere->write_comment( $worksheetLineSNPpereVsCNVmere,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'commentpLI'},x_scale => 2 );
				}	
				
				$worksheetLineSNPpereVsCNVmere ++;
				next;

			}	


			if( $hashFinalSortData{$rank}{$variant}{'worksheet'} =~ /SNPmCNVp/){

				$worksheetSNPmereVsCNVpere->write_row( $worksheetLineSNPmereVsCNVpere, 0, $hashFinalSortData{$rank}{$variant}{'finalArray'} );
				$worksheetSNPmereVsCNVpere->write_comment( $worksheetLineSNPmereVsCNVpere,$dicoColumnNbr{'MPA_ranking'}, $hashFinalSortData{$rank}{$variant}{'commentMPAscore'} ,x_scale => 2 );
				$worksheetSNPmereVsCNVpere->write_comment( $worksheetLineSNPmereVsCNVpere,$dicoColumnNbr{'gnomAD_genome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADgenome'} ,x_scale => 2 );
				$worksheetSNPmereVsCNVpere->write_comment( $worksheetLineSNPmereVsCNVpere,$dicoColumnNbr{'gnomAD_exome_ALL'}, $hashFinalSortData{$rank}{$variant}{'commentGnomADexome'} ,x_scale => 2 );
				$worksheetSNPmereVsCNVpere->write_comment( $worksheetLineSNPmereVsCNVpere,$dicoColumnNbr{'Genotype-'.$cas}, $hashFinalSortData{$rank}{$variant}{'commentGenotype'} ,x_scale => 2 );
				$worksheetSNPmereVsCNVpere->write_comment( $worksheetLineSNPmereVsCNVpere,$dicoColumnNbr{'Phenolyzer'}, $hashFinalSortData{$rank}{$variant}{'commentPhenolyzer'} ,x_scale => 2 );


				if ($hashFinalSortData{$rank}{$variant}{'commentpLI'} ne "."){
					$worksheetSNPmereVsCNVpere->write( $worksheetLineSNPmereVsCNVpere,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'finalArray'}[$dicoColumnNbr{'Gene.refGene'}]     ,$format_pLI );
					$worksheetSNPmereVsCNVpere->write_comment( $worksheetLineSNPmereVsCNVpere,$dicoColumnNbr{'Gene.refGene'},$hashFinalSortData{$rank}{$variant}{'commentpLI'},x_scale => 2 );
				}	
				$worksheetLineSNPmereVsCNVpere ++;
				
				next;

			}	








		}   # END IF TRIO
	
	
	
	
	
	
	}   # END FOREACH VARIANT
}	#END FOREACH RANK



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

print STDERR "Done!\n\n\n";

exit 0;
