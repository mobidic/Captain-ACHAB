#! /bin/bash


echo "usage : this_script.sh previousCustomVCF  newVCFtoCountnMerge.vcf " 

echo "it will output db.vcf"


echo "backup original customVCF" 
cp $1 "backup_"${1}



awk -F "\t" 'BEGIN{SAMPLES="";found=0;samplesFound="";HMZ=0;HTZ=0;TOTALSAMPLES=0;CHROMHEADER="";OFS="\t"}
{	if ($0 ~ /^##/  && SAMPLES==""){print;next};
	if ($0 ~ /^#CHROM/){
				CHROMHEADER= $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6"\t"$7"\t"$8"\t"$9;	

		split($0,sample,"\t")
		for (i=10;i<=NF; i++){
			if(i==10 && SAMPLES==""){
				SAMPLES=sample[i];
			}else{
				SAMPLES=SAMPLES"\t"sample[i];
			}		
			TOTALSAMPLES++;
		}
		CHROMHEADER = CHROMHEADER"\t"SAMPLES;
	}else{	if($8~/^found/){
			VARTAB[$1"_"$2"_"$4"_"$5]["full"]=$0;
			VARTAB[$1"_"$2"_"$4"_"$5]["startHalf"]=$1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6"\t"$7;
			VARTAB[$1"_"$2"_"$4"_"$5]["INFO"]=$8;
			VARTAB[$1"_"$2"_"$4"_"$5]["FORMAT"]=$9;
			VARTAB[$1"_"$2"_"$4"_"$5]["GENOTYPE"]="";
			for (i=10;i<=NF; i++){
				if(i==10 && SAMPLES!=""){
					VARTAB[$1"_"$2"_"$4"_"$5]["GENOTYPE"]=$i;
				}else{
					VARTAB[$1"_"$2"_"$4"_"$5]["GENOTYPE"]=VARTAB[$1"_"$2"_"$4"_"$5]["GENOTYPE"]"\t"$i;

				}
			}

			gsub("=|;","\t",$8);
			split($8,INFO,"\t");
			VARTAB[$1"_"$2"_"$4"_"$5]["found"]=INFO[2];
			VARTAB[$1"_"$2"_"$4"_"$5]["HTZ"]=INFO[4];
			VARTAB[$1"_"$2"_"$4"_"$5]["HMZ"]=INFO[6];
			VARTAB[$1"_"$2"_"$4"_"$5]["samplesFound"]=INFO[8];


		}else{
			
			for (i=10;i<=NF; i++){
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

			if ($1"_"$2"_"$4"_"$5 in VARTAB){

				VARTAB[$1"_"$2"_"$4"_"$5]["found"] += found;
				VARTAB[$1"_"$2"_"$4"_"$5]["HTZ"] += HTZ;
				VARTAB[$1"_"$2"_"$4"_"$5]["HMZ"] += HMZ;
				VARTAB[$1"_"$2"_"$4"_"$5]["samplesFound"] = VARTAB[$1"_"$2"_"$4"_"$5]["samplesFound"]"/"samplesFound;

				VARTAB[$1"_"$2"_"$4"_"$5]["INFO"]="found="VARTAB[$1"_"$2"_"$4"_"$5]["found"]";HTZ="VARTAB[$1"_"$2"_"$4"_"$5]["HTZ"]";HMZ="VARTAB[$1"_"$2"_"$4"_"$5]["HMZ"]";samplesFound="VARTAB[$1"_"$2"_"$4"_"$5]["samplesFound"];


			}else{
			
				$8="found="found";HTZ="HTZ";HMZ="HMZ";samplesFound="samplesFound";"$8;
				
				VARTAB[$1"_"$2"_"$4"_"$5]["full"]=$0;
				VARTAB[$1"_"$2"_"$4"_"$5]["startHalf"]=$1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6"\t"$7;
				VARTAB[$1"_"$2"_"$4"_"$5]["INFO"]=$8;
				VARTAB[$1"_"$2"_"$4"_"$5]["FORMAT"]=$9;
				VARTAB[$1"_"$2"_"$4"_"$5]["GENOTYPE"]="";
				for (i=10;i<=NF; i++){
					if(i==10 && SAMPLES!=""){
						VARTAB[$1"_"$2"_"$4"_"$5]["GENOTYPE"]=$i;
					}else{
						VARTAB[$1"_"$2"_"$4"_"$5]["GENOTYPE"]=VARTAB[$1"_"$2"_"$4"_"$5]["GENOTYPE"]"\t"$i;
    
					}
				}


			}


			found=0;
			samplesFound="";
			HMZ=0;
			HTZ=0;
		}
} 

}

END{print CHROMHEADER; for (var in VARTAB){print VARTAB[var]["startHalf"]"\t"VARTAB[var]["INFO"]"\t"VARTAB[var]["FORMAT"]"\t"VARTAB[var]["GENOTYPE"]  } }' $1 $2 > db.vcf


exit 0

