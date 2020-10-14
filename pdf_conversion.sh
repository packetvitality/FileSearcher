#!/bin/bash

#Attempts to convert all PDF files to text in the current directory.
for i in `ls | grep -i .pdf`; do 
	converted_file=`echo "$i" | cut -d"." -f1`; 
	pdftotext -layout "$i" "$converted_file".txt"; 
done
