# Python-Automation-Scripts

A Basic Automation File which allows word, excel and png documents to be converted to PDFs.

#After defining the import and from statements it consists of a Python OOP class for conversion. An object of the class Conversion needs to be instantiated to start using the function.

#Function 1 - list_files_idx
  This function uses the os walk method to look into the document which end with .idx and finds and replaces all instances of word, excel and png with pdf. Basically amending the index file.
  It does this by creating an absolute fake path and copying the index file there, and than it copies all the data of the original index file with the amendments and than moves the edited file over to the correct location.
  
#Function 2 - copy_files_idx
 This function basically changes the path of the original file location. Because all the files have to be moved from one folder location to another and the best way to do it was to just change the path and iterate using os.walk.
 
#Function 3 - convert_docx_pdf
 This function basically iterates over the folder and converts all the word, png and excel files to pdf.
 
#Function 4 - count_docs_completed
 This function basically counts all the separate documents in excel, word or png format and adds them to a set to ensure that all the files are individually added.
 
#Function 5 - count_pdfs
 This function counts all the pdfs in a folder
 
#Function 6 - count_docs
 This function counts all the documents in a folder
 
#Function 7 - get_completed_file_list
 This function gets all the file names from a text document after cleaning up the data in the each line and counts them.
 
#Function 8 - doc_as_docx
 This function basically uses win32 extension to convert word document in .doc extension to .docx extension.
 
#Function 9 - complete_docx_conv
 This function basically uses a text file and checks which files are in in, adds them to a list and iterates over the list to ensure that every document that needs to be converted from .doc to .dox is not in the file.
 
