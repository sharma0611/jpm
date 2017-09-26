#!/usr/bin/env python

from docxtpl import DocxTemplate
import os
from glob import glob
import comtypes.client
from PyPDF2 import PdfFileReader, PdfFileMerger
import sys
import csv

#setup 
config_dir = "./config"
output_dir_docs = "./docs"
output_dir_pdfs = "./pdfs"
output_pkgs = "./packages"
wdFormatPDF = 17

#create all required directories
os.makedirs(output_dir_docs, exist_ok=True)
os.makedirs(output_dir_pdfs, exist_ok=True)
os.makedirs(output_pkgs, exist_ok=True)

#define custom fn to grab filepaths in a directory that have certain extension
def get_filepaths(in_dir, file_extension):
    in_dir = in_dir.rstrip("/")
    results = [i for i in glob(in_dir + '/*.{}'.format(extension))]
    results = [os.path.abspath(i) for i in results]
    return results

#this function generates a tag to use as a title for a jop app to save as
def generate_tag(job_context):
    tag = job_context[job_context.keys()[0]]
    tag = tag.replace(" ", "")
    return tag
    
#stage config data for importing
# First, grab template docx
extension = "docx"
results = get_filepaths(config_dir, extension)
if len(results) != 1:
    print("Ensure you have exactly one template docx in /config")
    sys.exit(0)
else:
    template_fp = results[0]
# Next, grab data to be filled in
extension = "csv"
results = get_filepaths(config_dir, extension)
if len(results) != 1:
    print("Ensure you have exactly one csv in /config")
    sys.exit(0)
else:
    fill_data_fp = results[0]
# Finally, grab the pdfs that need to be added to the end of the pkg
extension = "pdf"
attach_pdfs = get_filepaths(config_dir, extension)

#Preprocessing
#Convert csv input fill data to list of dictionaries
with open(fill_data_fp) as f:
    jobs = [{k: str(v) for k, v in row.items()} for row in csv.DictReader(f, skipinitialspace=True)]

#Create all word documents from template
for job_context in jobs:
    #read in template
    curr_doc = DocxTemplate(template_fp)
    tag = generate_tag(job_context)
    curr_doc_name = "cover_letter_" + tag + ".docx"
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
