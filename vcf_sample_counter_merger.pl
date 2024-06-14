#!/usr/bin/perl

use strict;
use warnings;
use POSIX qw(strftime);

# Check if any arguments are provided
if (@ARGV == 0) {
	print "usage : this_script.pl outputFileName.vcf VCF1.vcf VCF2.vcf VCF3.vcf VCF4.vcf VCF5.vcf ...... \n";
	print "No arguments provided\n";
	exit 1;
}


my $output_file = shift @ARGV;
print "it will output in $output_file\n";
print "$#ARGV VCF files in input\n";

# Backup output file if it exists
my $date = strftime "%Y-%m-%d_at_%Hh%Mm%Ss", localtime;
if (-f $output_file) {
	my $backup_file = "$output_file.$date";
	system("cp", $output_file, $backup_file) == 0
		or die "Failed to backup file: $!";
}

# Initialize variables
my %VARTAB;
my $CHROMHEADER = "";
my $SAMPLES = "";
my $TOTALSAMPLES = 0;
my @sample;

open( OUT , ">", $output_file) or die("Cannot open output file ". $output_file);

foreach my $vcf (@ARGV){ 

	open( IN , "<$vcf") or die("Cannot open output file ". $vcf);
	print "Processing ".$vcf."\n"; 


	# Process each VCF file
	while (<IN>) {

		chomp;
		#get first header
		if (/^##/) {
			if ($SAMPLES eq "") {
				print OUT "$_\n";
			}
			next;
		}
		
		# get sample names and count
		if (/^#CHROM/) {
			$CHROMHEADER = join "\t", (split "\t")[0..7];
			
			@sample = split "\t";
			for (my $i = 9; $i <= $#sample; $i++) {
				
				if ($i == 9 && $SAMPLES eq ""){
					$SAMPLES = $sample[$i];
				}else{
					$SAMPLES .= "\t".$sample[$i];
				}
				$TOTALSAMPLES++;
			}
		} else {
			my @fields = split "\t";
			my $key = join "_", @fields[0, 1, 3, 4];

			$VARTAB{$key}{init} = $_;
			$VARTAB{$key}{startHalf} = join "\t", @fields[0..6];
			
			if ($fields[7] =~ /^found=/) {
				my %INFO = map { split /=/, $_ } split /;/, $fields[7];
				$VARTAB{$key}{found} = $INFO{found};
				$VARTAB{$key}{HTZ} = $INFO{HTZ};
				$VARTAB{$key}{HMZ} = $INFO{HMZ};
				$VARTAB{$key}{RECOMPUTED} = $INFO{RECOMPUTED};
				$VARTAB{$key}{samplesFoundHTZ} = $INFO{samplesFoundHTZ};
				$VARTAB{$key}{samplesFoundHMZ} = $INFO{samplesFoundHMZ};
				$SAMPLES = "DBfound";
			} else {
				my $found = 0;
				my $HTZ = 0;
				my $HMZ = 0;
				my $RECOMPUTED = 0;
				my $samplesFoundHTZ = "";
				my $samplesFoundHMZ = "";
				if (defined $VARTAB{$key}{samplesFoundHTZ} ){
					#nothing to do
				}else{
					$VARTAB{$key}{samplesFoundHTZ} = "";
				}
				
				if (defined $VARTAB{$key}{samplesFoundHMZ} ){
					#nothing to do
				}else{
					$VARTAB{$key}{samplesFoundHMZ} = "";
				}

				for (my $i = 9; $i <= $#fields; $i++) {
					my $allsamples = $VARTAB{$key}{samplesFoundHTZ}.$VARTAB{$key}{samplesFoundHMZ};
					my $pattern = $sample[$i]."/";
					#				print $sample[$i+1]."\n";

					if ($allsamples !~ /$pattern/) {
						if ($fields[$i] =~ /^0\/1:/ || $fields[$i] =~ /^1\/0:/ ) {
							$found++;
							$HTZ++;
							$samplesFoundHTZ .= $sample[$i]."/";
						} elsif ($fields[$i] =~ /^1\//) {
							$found++;
							$HMZ++;
							$samplesFoundHMZ .= $sample[$i]."/";
						} elsif ($fields[$i] =~ /^0\/0/) {
							my @recomputed = split /:/, $fields[$i];
							my @altdepth = split /,/, $recomputed[1];
							if ($altdepth[1] >= 1) {
								$found++;
								$RECOMPUTED++;
								$samplesFoundHTZ .= $sample[$i]."/";
							}
						}
					}
				}

				$VARTAB{$key}{found} += $found;
				$VARTAB{$key}{HTZ} += $HTZ;
				$VARTAB{$key}{HMZ} += $HMZ;
				$VARTAB{$key}{RECOMPUTED} += $RECOMPUTED;
				$VARTAB{$key}{samplesFoundHTZ} .= $samplesFoundHTZ;
				$VARTAB{$key}{samplesFoundHMZ} .= $samplesFoundHMZ;
			}
		}#END IF CHROM
	}#END WHILE
	close(IN);

}#END FOREACH vcf

# Print results
print "Prepare output\n"; 
print "Nb Samples processed: ".$TOTALSAMPLES."\n";

print OUT $CHROMHEADER."\t".$TOTALSAMPLES."_samples\n";
foreach my $key (sort keys %VARTAB) {
	if ( defined $VARTAB{$key}{samplesFoundHTZ} ){
	       	if( $VARTAB{$key}{samplesFoundHTZ} eq ""){
			$VARTAB{$key}{samplesFoundHTZ} = ".";
	 	}
	}else{
		 $VARTAB{$key}{samplesFoundHTZ} = ".";
	}
	if ( defined  $VARTAB{$key}{samplesFoundHMZ} ){
		if ($VARTAB{$key}{samplesFoundHMZ} eq ""){
			$VARTAB{$key}{samplesFoundHMZ} = ".";
	 	}
	}else{
		 $VARTAB{$key}{samplesFoundHMZ} = ".";
	}

	print OUT $VARTAB{$key}{startHalf}."\tfound=$VARTAB{$key}{found};HTZ=$VARTAB{$key}{HTZ};HMZ=$VARTAB{$key}{HMZ};RECOMPUTED=$VARTAB{$key}{RECOMPUTED};samplesFoundHTZ=$VARTAB{$key}{samplesFoundHTZ};samplesFoundHMZ=$VARTAB{$key}{samplesFoundHMZ}\n";
	
}

print "Done!\n";
exit 0;

