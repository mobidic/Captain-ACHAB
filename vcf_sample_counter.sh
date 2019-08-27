#! /bin/bash


echo "usage : this_script.sh   myVCF.vcf " 


echo "it will output db.vcf"

awk -F "\t" 'BEGIN{found=0;samplesFound="";HMZ=0;HTZ=0;OFS="\t"}
{if ($0 ~ /^#/){print;if($0 ~ /^#CHROM/){split($0,genotype,"\t")} 
	
	}else{for (i=10;i<=NF; i++){
			if($i ~ /^0\/1:/ ){
				found++;
				HTZ++;
				samplesFound=genotype[i]"/"samplesFound;
			}else if($i ~ /^1\// ){
				found++;
				HMZ++;
				samplesFound=genotype[i]"/"samplesFound;
			}
			
		}
		$8="found="found";HTZ="HTZ";HMZ="HMZ";samplesFound="samplesFound";"$8;
		print $0;
		found=0;
		samplesFound="";
		HMZ=0;
		HTZ=0;          
	} 

}' $1 > db.vcf


exit 0

