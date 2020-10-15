# FileSearcher
Python3 script to recursively search through unstructured directories and files.

You can simply point the script at a directory, and it will attempt to search the files using a provided keyword list. It should handle compressed files as well.

# Dependencies
Other then the requirements file, the script uses the python-magic which requires a seperate installer. See here:
https://pypi.org/project/python-magic/

# Pitfalls
PDF's are not handled very well. The 'PyPDF2' library used works well for the most part, but I've found on certain files it hangs (no error message or exit). I implemented some logic to skip the file if it takes too long. 
I prefer the option of converting the PDF's to text first using the linux 'pdftotext' tool. However, this is clunky to work into the script and involves another seperate install. 

# Future Work
I consider this a POC at this point. Currently investigating ways to increase reliability, increase performance, remove confusing dependencies, and make it easier to use. 

# Usage Example
Below is an example of searching through a directory:

    working_dir = "C:\\MyWorkingDirectory" # Base directory we are working from
    original_dir = "C:\\MyWorkingDirectory\\SearchDirectory" # Directory containing the files of interest
    keywords_file = "C:\\MyWorkingDirectory\\keywords.txt" # New Line Delimited File containing the keywords which will be searched for
    estimated_files = 1000000 # Used for more accurately displaying progress

    fs = FileSearcher(working_dir, original_dir, keywords_file, estimated_files=estimated_files)
    fs.create_dirs()
    searchme = original_dir 
    fs.process_directory(searchme)
    fs.cleanup_directories(searchme)
