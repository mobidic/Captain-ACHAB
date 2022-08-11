#! /bin/bash



if [ $# -eq 0 ]; then
	echo "usage : this_script.sh outputFileName.vcf  VCF1.vcf VCF2.vcf VCF3.vcf VCF4.vcf VCF5.vcf ...... " 
    echo "No arguments provided"
    exit 1
fi
		


echo "it will output in "$1  



echo "Backup outputFilename if file exists"  
#datestamp=`date +%F`
printf -v date '%(%Y-%m-%d_at_%Hh%Mm%Ss)T\n' -1
if [ -f "$1" ]; then
	cp $1 ${1}.$date
fi


awk -F "\t" 'BEGIN{SAMPLES="";found=0;samplesFound="";HMZ=0;HTZ=0;RECOMPUTED=0;TOTALSAMPLES=0;CHROMHEADER="";OFS="\t"}
{	if ($0 ~ /^##/){
		if(SAMPLES==""){print};next;
	}

	if ($0 ~ /^#CHROM/){
		CHROMHEADER= $1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6"\t"$7"\t"$8;	

		split($0,sample,"\t")
		for (i=10;i<=NF; i++){
			if(i==10 && SAMPLES==""){
				SAMPLES=sample[i];
			}else{
				SAMPLES=SAMPLES"\t"sample[i];
			}		
			TOTALSAMPLES++;
		}

	}else{ 
			VARTAB[$1"_"$2"_"$4"_"$5]["init"]=$0;
			VARTAB[$1"_"$2"_"$4"_"$5]["startHalf"]=$1"\t"$2"\t"$3"\t"$4"\t"$5"\t"$6"\t"$7;

			if($8~/^found=/){
				gsub("=|;","\t",$8);
				split($8,INFO,"\t");
				found=INFO[2];
				HTZ=INFO[4];
				HMZ=INFO[6];
				RECOMPUTED=INFO[8];
				samplesFound=INFO[10];
				SAMPLES="DBfound";
	
			}else{

				for (i=10;i<=NF; i++){
					allsamples = VARTAB[$1"_"$2"_"$4"_"$5]["samplesFound"];
					pattern = sample[i]"/";

					if (allsamples !~ pattern){
						if($i ~ /^0\/1:/ ){
							found++;
							HTZ++;
							samplesFound=sample[i]"/"samplesFound;
						}else if($i ~ /^1\// ){
							found++;
							HMZ++;
							samplesFound=sample[i]"/"samplesFound;
						}else if($i ~ /^0\/0/ ){
							split($i,recomputed,":")
							split(recomputed[2],altdepth,",")
							if (altdepth[2]>=1){
								found++;
								RECOMPUTED++;
								samplesFound=sample[i]"/"samplesFound;
							}
						}

					}
				}
			}


			VARTAB[$1"_"$2"_"$4"_"$5]["found"] += found;
			VARTAB[$1"_"$2"_"$4"_"$5]["HTZ"] += HTZ;
			VARTAB[$1"_"$2"_"$4"_"$5]["HMZ"] += HMZ;
			VARTAB[$1"_"$2"_"$4"_"$5]["RECOMPUTED"] += RECOMPUTED;
			VARTAB[$1"_"$2"_"$4"_"$5]["samplesFound"] = VARTAB[$1"_"$2"_"$4"_"$5]["samplesFound"]samplesFound;
  
			VARTAB[$1"_"$2"_"$4"_"$5]["INFO"]="found="VARTAB[$1"_"$2"_"$4"_"$5]["found"]";HTZ="VARTAB[$1"_"$2"_"$4"_"$5]["HTZ"]";HMZ="VARTAB[$1"_"$2"_"$4"_"$5]["HMZ"]";RECOMPUTED="VARTAB[$1"_"$2"_"$4"_"$5]["RECOMPUTED"]";samplesFound="VARTAB[$1"_"$2"_"$4"_"$5]["samplesFound"];
		
			found=0;
			samplesFound="";
			HMZ=0;
			HTZ=0;
			RECOMPUTED=0;
		}
} 

END{print CHROMHEADER; for (var in VARTAB){print VARTAB[var]["startHalf"]"\t"VARTAB[var]["INFO"] } }' $* > $1


exit 0

