# Captain ACHAB | Analysis Converter for Human who might Abhor Bioinformatics
--------------------------------------------------------------------------------
![JHI](image.png)

## Overview

Captain ACHAB is a simple and useful interface to analysis of WES data for molecular diagnosis.
This is the end of excel table with so much columns ! All necessary information is available in one look.

## Input 

A vcf annotated by ANNOVAR with MPA annotations and Phenolyzer predictions. 
See [MoBiDiC Prioritization Algorithm](https://github.com/mobidic/MPA/) and [Phenolyzer](https://github.com/WGLab/phenolyzer).

### Get custom annotations

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

According to use OMIM license, download the gene2map.txt at https://www.omim.org/downloads/

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

Cut the first column created by pandas and the gene_customfullxref.txt is ready to be use in ANNOVAR.

```bash
cut -f2- gene_customfullxref_tmp.txt > gene_customfullxref.txt
rm gene_customfullxref_tmp.txt 
```

### Annovar annotation 

A tutorial to install ANNOVAR and more informations are available at : [MoBiDiC Prioritization Algorithm](https://github.com/mobidic/MPA/)

Command line for vcf annotation by ANNOVAR with needed databases. 

```bash
perl path/to/table_annovar.pl path/to/example.vcf humandb/ -buildver hg19 -out path/to/output/name -remove -protocol refGene,refGene,clinvar_20170905,dbnsfp33a,spidex,dbscsnv11,gnomad_exome,gnomad_genome,intervar_20180118 -operation gx,g,f,f,f,f,f,f,f -nastring . -vcfinput -otherinfo -arg '-splicing 20','-hgvs',,,,,,, -xref example/gene_customfullxref.txt
```

### Phenolyzer annotation 

Tutorial to install Phenolyzer is available at [Phenolyzer](https://github.com/WGLab/phenolyzer). 

Installation 
```bash
git clone https://github.com/WGLab/phenolyzer
```

Command line 
```bash
perl disease_annotation.pl disease -f -p -ph -logistic -out disease/out
```

### MPA annotation

See installation and more informations about MPA at [MoBiDiC Prioritization Algorithm](https://github.com/mobidic/MPA/).

```bash
git clone https://github.com/mobidic/MPA.git
```

Command line for annotated vcf by ANNOVAR with MPA scores. 

```bash
python MPA.py -i name.hg19_multianno.vcf -o name.hg19_multianno_MPA.vcf
```

## Requirements

### Library

Python library : pandas and dependencies (only tested with python 2.7)
Perl library : BioPerl 
cpan ...
