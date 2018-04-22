# Captain ACHAB | Analysis Converter for Human who might Abhor Bioinformatics
--------------------------------------------------------------------------------
![JHI](achabjust.png)

## Overview

Captain ACHAB is a simple and useful interface to analysis of WES data for molecular diagnosis.
This is the end of excel table with so many columns ! All necessary information is available in one look.

## Input 

A vcf annotated by ANNOVAR with MPA annotations and Phenolyzer predictions. 
See [MoBiDiC Prioritization Algorithm](https://github.com/mobidic/MPA/) and [Phenolyzer](https://github.com/WGLab/phenolyzer).

### 1. Get custom annotations

To get unavailable annotations in ANNOVAR database into our vcf, we are going to add missense Z-score from ExAC and OMIM database into the gene_fullxref.txt from ANNOVAR available in the example folder.

#### Missense Z-score 

First download the database from ExAc (ftp.broadinstitute.org).

```bash
wget ftp://ftp.broadinstitute.org/pub/ExAC_release/release0.3.1/functional_gene_constraint/fordist_cleaned_exac_r03_march16_z_pli_rec_null_data.txt
```
Choose only columns neededs and reduce decimal to 3 numbers.

```bash
cut -f2,18 fordist_cleaned_exac_r03_march16_z_pli_rec_null_data.txt |  awk -F "\t" '{printf("%s,%.3f\n",$1,$2)}' | sed s/,/"\t"/g > missense_zscore.txt 
vim missense_zscore.txt ## change header "gene" to "#Gene_name" and name column 2 "Missense_Z_score" to allow recognition by pandas
```

#### OMIM 

According to use OMIM license, download the gene2map.txt at https://www.omim.org/downloads/.

```bash
tail -n+4 genemap2.txt | cut -f 9,13 > omim.tsv
vim omim.tsv ## change header "Approved Symbol" to "#Gene_name" to allow recognition by pandas
```

#### Merge with the gene_fullxref.txt

You need first to install pandas if needed.

```bash
pip install pandas
```
Use the merge function from pandas module to merge gene_fullxref.txt with OMIM and missense Z-score annotations. 

```python
import pandas

fullxref = pandas.read_table('gene_fullxref.txt') 
omim = pandas.read_table('omim.tsv')
zscore = pandas.read_table('missense_zscore.txt')

merge = pandas.merge(fullxref,omim, on="#Gene_name", how="left", left_index=True)
mergeFinal = pandas.merge(merge,zscore, on="#Gene_name", how="left", left_index=True)

mergeFinal.to_csv('gene_customfullxref_tmp.txt',sep='\t')
```

Cut the first column created by pandas and the gene_customfullxref.txt is ready to be use in ANNOVAR. sed and awk command are used to replace some characters not compatible for regex in ANNOVAR. Last awk command fill the empty cells with a point.

```bash
cut -f2- gene_customfullxref_tmp.txt | sed 's/+/plus/g' | awk 'BEGIN{FS=OFS="\t"} {for (i=6;i<=7;i++) gsub(/-/,"_",$i)}1' |  awk 'BEGIN{FS=OFS="\t"} {gsub(/-/,"_",$(NF-1))}1' | sed 's/(congenital with brain and eye anomalies,/(congenital with brain and eye anomalies),/g' | awk -F"\t" -v OFS="\t" '{for (i=1;i<=NF;i++) {if ($i == "") $i="."} print $0}' > gene_customfullxref.txt
rm gene_customfullxref_tmp.txt 
```

### 2. Annovar annotation 

A tutorial to install ANNOVAR and more informations are available at : [MoBiDiC Prioritization Algorithm](https://github.com/mobidic/MPA/)

Note: multiallelic lines from vcf have to be split before annotation ( using: sort vcf then
bcftools-1.3.1/htslib-1.3.1/bgzip -i example.sort.vcf
bcftools-1.3.1/bcftools norm -O v -m - -o example.norm.vcf example.sort.vcf.gz)

Command line for vcf annotation by ANNOVAR with needed databases. 

```bash
perl path/to/table_annovar.pl path/to/example.vcf humandb/ -buildver hg19 -out path/to/output/name -remove -protocol refGene,refGene,clinvar_20170905,dbnsfp33a,spidex,dbscsnv11,gnomad_exome,gnomad_genome,intervar_20180118 -operation gx,g,f,f,f,f,f,f,f -nastring . -vcfinput -otherinfo -arg '-splicing 20','-hgvs',,,,,,, -xref example/gene_customfullxref.txt
```

### 3. MPA annotation

See installation and more informations about MPA at [MoBiDiC Prioritization Algorithm](https://github.com/mobidic/MPA/).

```bash
git clone https://github.com/mobidic/MPA.git
```

Command line for annotated vcf by ANNOVAR with MPA scores.

```bash
python MPA.py -i name.hg19_multianno.vcf -o name.hg19_multianno_MPA.vcf
```

### 4. Phenolyzer annotation 

Tutorial to install Phenolyzer is available at [Phenolyzer](https://github.com/WGLab/phenolyzer). 

Installation (need Bioperl and Graph, easy to install with cpanm)
```bash
git clone https://github.com/WGLab/phenolyzer
```

Create a disease file where you can add your HPO phenotypes (one line per phenotype).

```bash
vim disease.txt
```

Command line to get predictions for Phenolyzer and the out.predicted_gene_scores.

```bash
perl disease_annotation.pl disease.txt -f -p -ph -logistic -out disease/out
```

## Captain ACHAB Command

Installation (need Switch, Excel::Writer::XLSX, easy to install with cpanm)

```bash
https://github.com/mobidic/Captain-ACHAB.git
```

Command line to use Captain ACHAB 

```
# USAGE : perl achab.pl --vcf <vcf_file> --case <index_sample_name> --dad <father_sample_name> --mum <mother_sample_name> --control <control_sample_name>  --trio <YES|NO> --candidats <file with gene symbol of interest>  --phenolyzerFile <phenolyzer output file suffixed by predicted_gene_scores>   --popFreqThr <allelic frequency threshold from 0 to 1 default=0.01>  --customInfo  <info name (will be added in a new column)>
```

## Requirements

### Library

Python library : pandas and dependencies (only tested with python 2.7)

Perl library via cpanm : BioPerl, Graph, Switch, Excel::Writer::XLSX


