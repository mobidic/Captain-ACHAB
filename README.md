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

To get unavailable annotations in ANNOVAR database into our vcf, we are going to add missense Z-score from ExAC and OMIM database into the gene_fullxref.txt 

### Annovar annotation 

Command line 

```bash
perl path/to/table_annovar.pl path/to/example.vcf humandb/ -buildver hg19 -out path/to/output/name -remove -protocol refGene,refGene,clinvar_20170130,dbnsfp33a,spidex,dbscsnv11,gnomad_exome,gnomad_genome -operation gx,g,f,f,f,f,f,f -nastring . -vcfinput -otherinfo -arg '-splicing 20','-hgvs',,,,,, -xref example/gene_fullxref.txt
```

## Requirements

### Library

Perl library : cpan ...
