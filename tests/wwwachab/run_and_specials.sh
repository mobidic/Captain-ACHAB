#!/usr/bin/env bash


### Bash script to Achab


# WARN: Clean PATH from custom stuffs:
#       (similar to when somebody else is running pipeline)
OLD_PATH=$PATH && PATH="/usr/local/bin:/usr/bin:/usr/local/sbin:/usr/sbin"  # Minimal PATH

# Activate Conda env with all required dependencies:
# WARN: Use 'Exome' one
source /etc/profile.d/conda.sh && conda activate /mnt/Bioinfo/Softs/src/conda/envs/Exome_prod


set -euo pipefail  # Bash dev best practice

set -x  # DEBUG


# Get params:
params_file=$1  # List of params that will be 'catted'
csvtkExe=/mnt/Bioinfo/Softs/bin/csvtk


# Create out-dir:
achabOutDir=./out-Achab
mkdir "$achabOutDir"


# Run Achab:
$CONDA_PREFIX/bin/perl wwwachab.pl $(cat $params_file)


# Find produced xlsx (can be either newHope):
foundAchab=$(find "$achabOutDir" -type f -name "*xlsx" ! -name "*_poorCoverage.xlsx")

if [ -z "$foundAchab" ] ; then
	echo "WARN: NO 'Achab.xlsx' found -> '--version' mode ?" >> /dev/stderr
	exit
fi


# Run 'specials' (= extract each sheet to check its content):
testOutDir=./out-test
mkdir "$testOutDir"

# File listing sheets:
listSheetsOut="$testOutDir"/achab_sheets.tsv
"$csvtkExe" xlsx2csv --list-sheets "$foundAchab" > "$listSheetsOut"

for a_sheet in $("$csvtkExe" cut -t -f sheet "$listSheetsOut" | awk 'NR>1') ; do
	"$csvtkExe" xlsx2csv --sheet-name "$a_sheet" -o "$testOutDir"/sheet_"$a_sheet".csv "$foundAchab"
done

# Checksum each extracted sheets file:
md5sum "$testOutDir"/sheet_*.csv > "$testOutDir"/sheetS.md5


# Run 'specials I' (= extract only 'variants' ie. 'CHR-POS-REF-ALT' from each sheet, preserve ORDER)
sortedOutDir=./out-ranked
mkdir --verbose "$sortedOutDir"

for a_sheet in $("$csvtkExe" cut -t -f sheet "$listSheetsOut" | awk 'NR>1') ; do
	if grep --quiet --word-regexp "#CHROMPOSREFALT" <("$csvtkExe" headers "$testOutDir"/sheet_"$a_sheet".csv) ; then  # Run only on sheets having a "#CHROMPOSREFALT" column
		"$csvtkExe" cut --fields "#CHROMPOSREFALT" "$testOutDir"/sheet_"$a_sheet".csv |
			"$csvtkExe" del-header > "$sortedOutDir"/vars_"$a_sheet".csv
	fi
done

# Checksum each extracted variants file:
md5sum "$sortedOutDir"/vars_*.csv > "$sortedOutDir"/variantS.md5


# Run 'specials II' (= same as 'specials I' but re-sort variants list)
REsortedOutDir=./out-REsorted
mkdir --verbose "$REsortedOutDir"

for a_sheet in $("$csvtkExe" cut -t -f sheet "$listSheetsOut" | awk 'NR>1') ; do
	if grep --quiet --word-regexp "#CHROMPOSREFALT" <("$csvtkExe" headers "$testOutDir"/sheet_"$a_sheet".csv) ; then  # Run only on sheets having a "#CHROMPOSREFALT" column
		"$csvtkExe" cut --fields "#CHROMPOSREFALT" "$testOutDir"/sheet_"$a_sheet".csv |
			"$csvtkExe" del-header |
			LC_ALL=C sort > "$REsortedOutDir"/vars_"$a_sheet".csv
	fi
done

# Checksum each extracted variants file:
md5sum "$REsortedOutDir"/vars_*.csv > "$REsortedOutDir"/variantS.md5
