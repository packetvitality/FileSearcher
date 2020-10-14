# FileSearcher
Python3 script to search through unstructured directories and files.

You can simply point the script at a directories, and it will attempt to search the files using a provided keyword list. It should handle compressed files as well.

# Dependencies
Other then the requirements file, the script uses the python-magic which requires a seperate installer. See here:
https://pypi.org/project/python-magic/

# Pitfalls
PDF's are not handled very well. The 'PyPDF2' library uses mostly works well, but I've found on certain files it hangs (no error message or exit).
I prefer the option of converting the PDF's to text first using the linux 'pdftotext' tool. However, this is clunky to work into the script and involves another seperate install. 

# Future Work
I consider this a POC at this point. Currently investigating ways to increase reliability, increase performance, remove confusing dependencies, and make it easier to use. 
