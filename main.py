#!/usr/bin/env python

from docxtpl import DocxTemplate
from config.joblistings import jobs
import os
from glob import glob
import comtypes.client
from PyPDF2 import PdfFileReader, PdfFileMerger

#setup 
template_file = "./config/cover_letter_template.docx"
output_dir_docs = "./docs"
output_dir_pdfs = "./pdfs"
output_pkgs = "./packages"
include_dir = "./include"
wdFormatPDF = 17

#create all required directories
os.makedirs(output_dir_docs, exist_ok=True)
os.makedirs(output_dir_pdfs, exist_ok=True)
os.makedirs(output_pkgs, exist_ok=True)

#Create all word documents from template
for job in jobs:
    #read in template
    curr_doc = DocxTemplate(template_file)
    curr_doc_name = "cover_letter_" + job.replace(" ", "") + ".docx"
    curr_doc_path = output_dir_docs + "/" + curr_doc_name
    context = {"company_name" : job}
    #replace contents
    curr_doc.render(context)
    #save to output docs
    if os.path.isfile(curr_doc_path):
        os.remove(curr_doc_path)
    curr_doc.save(curr_doc_path)

#Get all word documents, convert to pdfs
word = comtypes.client.CreateObject('Word.Application')

for word_doc in glob(output_dir_docs + "/*.docx"):
    basename = os.path.basename(word_doc)
    in_file = os.path.abspath(word_doc)
    if '~' in in_file:
        continue
    out_file = output_dir_pdfs + "/" + basename[:-5] + ".pdf" # [:-5] removes .docx extension
    out_file = os.path.abspath(out_file)
    print(in_file)
    doc = word.Documents.Open(in_file)
    if os.path.isfile(out_file):
        os.remove(out_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()

word.Quit()


#Get all pdf documents, append pages needed
add_files = glob(include_dir + "/*.pdf")

for pdf_doc in glob(output_dir_pdfs + "/*.pdf"):
    basename = os.path.basename(pdf_doc)
    pkg_output_path = output_pkgs + "/" + basename
    #initiate the package
    output = PdfFileMerger()
    #add the cover letter
    with open(pdf_doc, 'rb') as cover_letter:
        output.append(PdfFileReader(cover_letter))
    #add each file from includes
    for add_file in add_files:
        with open(add_file, 'rb') as temp:
            output.append(PdfFileReader(temp))
    if os.path.isfile(pkg_output_path + ".pdf"):
        os.remove(pkg_output_path + ".pdf")
    #Now output the package with correct filename
    output.write(pkg_output_path)
