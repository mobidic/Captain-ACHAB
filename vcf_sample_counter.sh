#! /bin/bash


echo "usage : this_script.sh   myVCF.vcf " 


echo "it will output db.vcf"


awk 'BEGIN{found=0;samplesFound="";OFS="\t"}
{if ($0 ~ /^#/){print;if($0 ~ /^#CHROM/){split($0,genotype,"\t") }
	
	}else{for (i=10;i<=NF; i++){
			if($i ~ /^0\/1:/ || $i ~ /^1\// ){
				found++;samplesFound=genotype[i]"/"samplesFound};
			};
		$8="found="found";samplesFound="samplesFound";"$8; print $0;found=0;samplesFound=""          
	} 

}' $1 > db.vcf


exit 0

