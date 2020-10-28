# -*- coding: utf-8 -*-
"""
Final Version

Created on Fri Oct 16 12:59:48 2020

@author: umaraftab
"""

from fpdf import FPDF
from os import listdir
import os
from pathlib import Path
from PIL import Image
from docx2pdf import convert as conv
from win32com import client
import win32api
from subprocess import check_call
from distutils.dir_util import copy_tree
import comtypes.client
import re
from glob import glob
import win32com.client as win32
from win32com.client import constants
from tempfile import mkstemp
from shutil import move, copymode
from os import fdopen, remove

class Conversion:
    def __init__(self):
        pass
    
    
    def list_files_idx(self,startpath):
      for root, dirs, files in os.walk(startpath):
          for f in files:
              if f.lower().endswith(('.idx')):
                  filepath = Path(os.path.join(root,f))
                  #print(filepath)
                  fh, abs_path = mkstemp()
                  with fdopen(fh,'w') as new_file:
                      with open(filepath, "r") as fileItem:
                          for item in fileItem:
                              #print(item)
                              #new_file.write(item)
                            if item.find('.docx') != -1:
                                new_file.write(item.replace('.docx','.pdf'))
                            elif item.find('.xlsx') != -1:
                                new_file.write(item.replace('.xlsx','.pdf'))
                            elif item.find('.png') != -1:
                                new_file.write(item.replace('.png','.pdf'))
                            elif item.find('.doc') != -1:
                                new_file.write(item.replace('.doc','.pdf'))
                            elif item.find('.xls') != -1:
                                new_file.write(item.replace('.xls','.pdf'))
                            elif item.find('.pdfx') !=-1:
                                new_file.write(item.replace('.pdfx','.pdf'))
                            elif item.find('.pdf') == 1:
                                print(item)
                  copymode(filepath, abs_path)
    #Remove original file
                  remove(filepath)
    #Move new file
                  move(abs_path, filepath)
                  
    def copy_files_idx(self,startpath,str1,str2):
      for root, dirs, files in os.walk(startpath):
          for f in files:
              if f.lower().endswith(('.idx')):
                  original_path = Path(os.path.join(root,f))
                  print(original_path)
                  new_path =os.path.normpath(str(original_path).replace(str1,str2))
                  print(new_path)
    #Remove original file
                  move(original_path, new_path)
                  
    def convert_docx_pdf(self,startpath,fileList):
        for root, dirs, files in os.walk(startpath):
            level = root.replace(startpath, '').count(os.sep)
            indent = ' ' * 4 * (level)
            print('{}{}/'.format(indent, os.path.basename(root)))
            subindent = ' ' * 4 * (level + 1)
            for f in files:
                if f.lower().endswith(('.docx')) and f not in fileList:
                    print(f)
                    #If the file is a word document, firstly a new .pdf extension is created to convert and the doc2xpdf library for conversion to pdf is used
                    if f.lower().endswith(('.docx')):
                        new_ext_doc='.pdf'
                        filepathdoc = Path(os.path.join(root,f))
                        pdf_filepathdoc = str(filepathdoc).replace("".join(filepathdoc.suffixes),new_ext_doc)
                        conv(filepathdoc)
                        conv(filepathdoc, pdf_filepathdoc)
                    
                    elif f.lower().endswith(('.png')):
                        new_ext='.pdf'
                        filepath = Path(os.path.join(root,f))
                        image1= Image.open(filepath)
                        im1=image1.convert('RGB')
                        pdf_filepath = str(filepath).replace("".join(filepath.suffixes),new_ext)
                        im1.save(pdf_filepath)
  
                    #If the file is an excel, firstly a new .pdf extension is created to convert and the win32 library for conversion to pdf is used
                    elif f.lower().endswith(('.xlsx')) or f.lower().endswith(('.xls')):
                        new_ext_xl='.pdf'
                        filepathxl = Path(os.path.join(root,f))
                        pdf_filepathxl = str(filepathxl).replace("".join(filepathxl.suffixes),new_ext_xl)
    
                        input_file = filepathxl
                        #give your file name with valid path 
                        output_file = pdf_filepathxl
                        #give valid output file name and path
                        app = client.DispatchEx("Excel.Application")
                        app.Interactive = False
                        app.Visible = False
                        Workbook = app.Workbooks.Open(input_file)
                        try:
                            Workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
                        except Exception as e:
                            print("Failed to convert in PDF format.Please confirm environment meets all the requirements  and try again")
                            print(str(e))
                        finally:
                            Workbook.Close()
                            
    def count_docs_completed(self):
        count=0
        countExcel=0
        excelSet=set()
        countWord=0
        countPng=0
        pngSet=set()
        wordSet=set()
        fileList=[]
        with open(r'filesDone.txt','r') as f:
            listWord='100%'
            listNum='1/1'
            for line in f:
                if listWord in line or listNum in line:
                    pass
                else:
                    line=str(line)
                    if re.search(r'.doc$', line):
                        pass
                    elif re.search(r'^~',line):
                        pass
                    else:
                        count=count+1
                        if re.search(r'.xls$',line) or re.search(r'.xlsx$',line):
                            countExcel=countExcel+1
                            excelSet.add(line)
                        elif re.search(r'.docx$',line):
                            countWord=countWord+1
                            wordSet.add(line)
                        elif re.search(r'.png',line):
                            countPng=countPng+1
                            pngSet.add(line)
                        fileItem=str(line)
                        fileItem=fileItem.strip().split("\n")
                        fileList.append(fileItem[0])
        completed_set = wordSet.union(pngSet).union(excelSet)                
        return completed_set

    def count_pdf(self,startpath):
        pdf_count=0
        pdf_set=set()
        for root, dirs, files in os.walk(startpath):
            for f in files:
                 if f.lower().endswith(('.pdf')):
                        pdf_count=pdf_count+1
                        pdf_set.add(f)
        return pdf_count,pdf_set
    
    def count_docs(self,startpath):
        doc_count=0
        doc_set=set()
        for root, dirs, files in os.walk(startpath):
            for f in files:
                 if f.lower().endswith(('.xlsx','.png','.docx','.xls')): 
                        doc_count=doc_count+1
                        doc_set.add(f)
        return doc_count,doc_set
    
    def get_completed_file_list(self,filepath):
        count=0
        fileList=[]
        with open(filepath,'r') as f:
            listWord='100%'
            listNum='1/1'
            for line in f:
                if listWord in line or listNum in line:
                    pass
                else:
                    line=str(line)
                    if re.search(r'.doc$', line):
                        pass
                    else:
                        count=count+1
                        fileItem=str(line)
                        fileItem=fileItem.strip().split("\n")
                        fileList.append(fileItem[0])
        return fileList
    
    def doc_as_docx(self,path):
        # Opening MS Word
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(path)
        doc.Activate ()
    
        # Rename path with .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
    
        # Save and Close
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        doc.Close(False)
        
    def complete_docx_conv(self,fileLink,startpath):
        link=fileLink
        listDoc=[]
        with open(link,'r') as f:
            listDoc.append(str(f.read()))
        listDoc=[x.strip().split("\n") for x in listDoc]
        
        paths = glob(startpath+'\**\*.doc', recursive=True)
        for path in paths:
            if path in listDoc[0]:
                pass
            else:
                print(startpath)
                self.doc_as_docx(startpath)
                
conv = Conversion()


pdf_count,pdf_set=conv.count_pdf()
doc_count_OTTR,doc_set_OTTR=conv.count_docs('')
doc_count_ARCH,doc_set_ARCH=conv.count_docs()
completed_set=conv.count_docs_completed()


document_set_txt = set(map(lambda x:x.split(".")[0],completed_set))
pdf_document_set = set(map(lambda x:x.split(".")[0],pdf_set))
document_set_OTTR = set(map(lambda x:x.split(".")[0],doc_set_OTTR))

print("The length of pdf Set is",len(pdf_document_set),pdf_count)
print("The length of Document Set TXT",len(document_set_txt))


print("Documents in OTTR1 and OTTR_Archive")
print("The documents in OTTR1 :",len(doc_set_OTTR),doc_count_OTTR)
print("The documents in OTTR_ARCHIVE :",len(doc_set_ARCH),doc_count_ARCH)

difference_set=document_set_txt.difference(pdf_document_set)

print("The different files based on txt",len(difference_set))
for item in difference_set:
    print(item+".docx")
    

difference_set_OTTR=document_set_OTTR.difference(pdf_document_set)

print("The different files based on txt",len(difference_set_OTTR))
for item in difference_set_OTTR:
    print(item)
