# Captain ACHAB | Analysis Converter for Human who might Abhor Bioinformatics
--------------------------------------------------------------------------------
![JHI](image.png)

## Overview

Captain ACHAB is a simple and useful interface to analysis of WES data for molecular diagnosis.
This is the end of excel table with so much columns ! All necessary information is available in one look.

## Input 

A vcf annotated by ANNOVAR with MPA annotations. 
See [MoBiDiC Prioritization Algorithm](https://github.com/mobidic/MPA/).

### Get custom annotations

To get unavailable annotations in ANNOVAR database into our vcf, we are going to add missense Z-score from ExAC and OMIM database into the gene_fullxref.txt from ANNOVAR.

#### Missense Z-score 

First download the database from ExAc (ftp.broadinstitute.org).

```bash
wget ftp://ftp.broadinstitute.org/pub/ExAC_release/release0.3.1/functional_gene_constraint/fordist_cleaned_exac_r03_march16_z_pli_rec_null_data.txt
```
Choose only columns neededs

```bash
cut -f2,18 fordist_cleaned_exac_r03_march16_z_pli_rec_null_data.txt > missense_zscore.txt
vim missense_zscore.txt ## change header "gene" to "#Gene_name" to allow recognition by pandas
```

#### OMIM 

According to use OMIM license, download the gene2map.txt at https://www.omim.org/downloads/

```bash

```

### Annovar annotation 

Command line 

```bash
perl path/to/table_annovar.pl path/to/example.vcf humandb/ -buildver hg19 -out path/to/output/name -remove -protocol refGene,refGene,clinvar_20170130,dbnsfp33a,spidex,dbscsnv11,gnomad_exome,gnomad_genome -operation gx,g,f,f,f,f,f,f -nastring . -vcfinput -otherinfo -arg '-splicing 20','-hgvs',,,,,, -xref example/gene_fullxref.txt
```

## Requirements

### Library

Perl library : cpan ...
