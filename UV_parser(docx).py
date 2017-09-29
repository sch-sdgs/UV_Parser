import os
import sys
import re
import csv
import zipfile
import argparse
import docx2txt
import fileinput
from datetime import datetime
import glob
from itertools import islice
import numpy as np
import pandas as pd
import unicodedata 
import xml.etree.ElementTree as ET       #import of diffrent modules needed by the code    
from dateutil.parser import parse
from datetime import datetime
nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}    #giving acces to xml tags 

def qn(tag): #reformating the xml tags 

    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{{{}}}{}'.format(uri, tagroot)

def xml2text(xml): #parses the xml docx  # Parser based on qn tags  p = paragraph # t=text # tab= table # cr = cell # br = row # takes out the text in the tags and strips it and saves it as multi-lined string
    text = ''
    root = ET.fromstring(xml)
    for child in root.iter():
      if child.tag == qn('w:t'):
        t_text = child.text
        text += t_text
      elif child.tag == (qn('w:tab')):
        text += '\t'
      elif child.tag in (qn('w:br'), qn('w:cr')):  
        text += '\n'
      elif child.tag == (qn('w:p')):
        text += '\n\n' 
            

    return text.strip() 
#This is checking if the file is a uv form based on a matching phrase
def check_file(filenames):
  #test1 = "/home/bioinfo/anakev/wc/UVs_fromblobfiles/Test2/"
    for filename in glob.glob(os.path.join(filenames, '*.docx')):
      try:
        print filename
        document = docx2txt.process(filename)
        match = "Unclassified Variant Evaluation Form (UV Form)"
    #match1 = "Document reference number:  402.002"
    #match2 = "Document reference number:  602.009"
        if match in document:
          process(filename)
    #elif match in document and match2 in document:
        else:
          f = open('log.txt', 'ar')
          f.write("This is not a UV form" + filename + "\n")
          f.close()
          print "This is not a UV form" + filename          
      except Exception, e:
        f = open('log.txt', 'ar')
        f.write('An exceptional thing happed - %s' % e + docx + "\n")
        f.close()
        print 'process error this' + docx + ' is not parsable '       
        pass
def process(docx, img_dir=None):   #main method 
  try:
    zipf = zipfile.ZipFile(docx)                            #opening the docx files which are zip-formated
    doc_xml = 'word/document.xml'
    text = xml2text(zipf.read(doc_xml))                     # takes out and saves the parsed text
    text1 = text.splitlines()                               # splits the multi line sting into strings
    text2 = filter(None , text1)                            
    for anything, i in enumerate (text2):                   # this is making the start point for a list
      if i == 'Disease':
        dis = anything 
    for ref_seq, j in enumerate (text2):
      if j == 'Ref Seq No':
        ref = ref_seq ################################# this on is 0 
    for gene, j in enumerate (text2):
      if j == 'Gene':
        gen = gene
    for cdna, j in enumerate (text2):
      if j == 'cDNA':
        cdn = cdna
    for dna, j in enumerate (text2):
      if j == 'DNA':
        dn = dna
    for exon, j in enumerate (text2):
      if j == 'Exon':
        exo = exon
    for cprotein, j in enumerate (text2):
      if j == 'Protein':
        cprotei = cprotein
    for classificaiton, j in enumerate (text2):
      if j == 'Classification':
        classificaito = classificaiton                       
    for conclusion, j in enumerate (text2):
      if j == 'CONCLUSION':
        conc = conclusion
        #print sanity check  
    for familydetails, j in enumerate (text2):
      if j == 'Family details':
        fam = familydetails
    for pedigree_number, p in enumerate (text2): 
      if p in('Pedigree Number','Pedigree number'):
        Pedigree = pedigree_number
    final = ' '.join(text2[conc+1:fam]).replace(",", " ") # this is setting up the list, start/end point removing , and joining list items  
    for summary, j in enumerate (text2):
      if j == 'SUMMARY':
        summ = summary
    for extent_of, j in enumerate (text2):
      if j == '1.  Extent of Analysis ':
        ext = extent_of
    for Idanddate, j in enumerate (text2):
      if j == 'Completed by':
        inf = Idanddate
    for Checked_by, j in enumerate (text2):
      if j == 'Checked by' : 
        chec = Checked_by 
    for LabNumber, j in enumerate (text2):
      if j == 'Lab Number':
        lab = LabNumber        
    disease_only_list  = text2[dis+1:ref]
    disease_only_list1 = []
    disease_only_list1.extend(disease_only_list) #this is appending list items to each other
    disease = ', '.join(disease_only_list1)
    disease1 =  disease.replace(",", " ")
    gene_only_list  = text2[gen+1:cdn]
    gene_only_list1 = []
    gene_only_list1.extend(gene_only_list)
    gene = ', '.join(gene_only_list1)
    gene1 =  gene.replace(",", " ")
    exon_only_list  = text2[exo+1:cprotei-2]
    exon_only_list1 = []
    exon_only_list1.extend(exon_only_list)
    exon = ', '.join(exon_only_list1)
    exon1 =  exon.replace(",", " ")
    cdna_transcript_only_list  = text2[cdn+1:dn]
    cdna_transcript_only_list1 = []
    cdna_transcript_only_list1.extend(cdna_transcript_only_list)
    cdna_transcript  = ', '.join(cdna_transcript_only_list1)   
    cdna_transcript1 = cdna_transcript.replace(",", " ")
    cdna_variant_only_list  = text2[dn+1:exo]
    cdna_variant_list1 = []
    cdna_variant_list1.extend(cdna_variant_only_list)
    cdna_variant  = ', '.join(cdna_variant_list1)
    cdna_variant1 = cdna_variant.replace(",", " ")
    protein_variant_only_list  = text2[exo+3:cprotei]
    protein_variant_only_list1 = []
    protein_variant_only_list1.extend(protein_variant_only_list)
    protein_variant  = ' '.join(protein_variant_only_list1)
    protein_change_only_list  = text2[cprotei+1:classificaito]   
    protein_change_only_list1 = []
    protein_change_only_list1.extend(protein_change_only_list)
    protein_change  = ' '.join(protein_change_only_list1)
    protein_transcript = protein_change.replace(",", " ")
    final_classification  = text2[classificaito:conc]
    final_classification1 = []
    final_classification1.extend(final_classification)
    final_class = final_classification1[2].replace(",", " ")
    test = text2[17].replace(",", " ")  
    final_1 = text2[summ:inf]
    summary_list = []
    summary_list.extend(final_1)  
    try: #this tries to parse the date if it doesn' it prints out the whole list. 
      finaldate = text2[inf+2:chec]
      final_date = []
      final_date.extend(finaldate)
      listsomething1 = ' '.join(final_date).encode("utf-8")
      date_by = parse(listsomething1)
      Date_by = str(date_by).replace(",", " ")
    except:
      finaldate1 = text2[inf:chec]    #+2
      final_date = []
      final_date.extend(finaldate1)
      listsomething1= ' '.join(final_date)
      Date_by = smart_str(listsomething1).replace(",", " ")    #.encode("utf-8")
    try:
      finaldate2 = text2[chec+2:ext]
      final_date = []
      final_date.extend(finaldate2)
      listsomething1 = ' '.join(final_date).encode("utf-8")
      date_by = parse(listsomething1)
      Date_by1 = str(date_by).replace(",", " ")
    except:
      finaldate3 = text2[chec:ext]    #+2
      final_date = []
      final_date.extend(finaldate3)
      listsomething1= ' '.join(final_date)
      Date_by1 = smart_str(listsomething1).replace(",", " ")   #.encode("utf-8")
    final_datetime = text2[inf:ext]
    final_date = []
    final_date.extend(final_datetime)
    print Date_by
    print Date_by1 # sanity checks 
    # assigning the var's to
    p_number = text2[fam:summ]
    p_number1 = []
    p_number1.extend(p_number)
    pedigree_number1 = p_number1[2].replace(",", " ")
    Index_Case = p_number1[4].replace(",", " ")
    Lab_Number = p_number1[6].replace(",", " ")
    Completed_by = final_date[1].replace(",", " ")
    Checked_by2 = final_date[4].replace(",", " ")
    Extent_of_Analysis = summary_list[4].replace(",", " ")
    Population_Frequency = summary_list[7].replace(",", " ")
    Web_based_Literature_Search = summary_list[10].replace(",", " ")
    Frameshift_Mutation = summary_list[13].replace(",", " ")
    In_frame_Deletion_or_Insertion = summary_list[16].replace(",", " ")
    Nonsense_Mutation = summary_list[19].replace(",", " ")
    Evolutionary_Conservation = summary_list[22].replace(",", " ")
    Severity_of_Amino_Acid_Substitution = summary_list[25].replace(",", " ")
    Splice_Site_Prediction = summary_list[28].replace(",", " ")
    Functional_Studies = summary_list[31].replace(",", " ")
    mRNA_Analysis = summary_list[34].replace(",", " ")
    Genotype_Phenotype_Correlation = summary_list[37].replace(",", " ")
    Co_Segregation = summary_list[40].replace(",", " ") 
    De_Novo_Variant = summary_list[43].replace(",", " ")
    Phase_established_for_Multiple_Variants_in_cis_or_in_trans = summary_list[46]
    Tumour_Studies = summary_list[49].replace(",", " ")
    Heteroplasmy_and_Mitochondrial_DNA_changes = summary_list[52]
    final_record = smart_str(docx).replace(","," ") + "," + smart_str(disease1) + "," + smart_str(gene1) + ","+ smart_str(exon1) + "," + smart_str(cdna_transcript1) + "," + smart_str(cdna_variant1) + "," + smart_str(protein_transcript).replace(","," ") + "," + smart_str(protein_change).replace(","," ") + "," + smart_str(final_class) + "," + smart_str(pedigree_number1) +"," + smart_str(Completed_by) + "," + Date_by + "," + smart_str(Checked_by2) + "," + Date_by1 + "," + smart_str(Index_Case) + ","+ smart_str(Lab_Number) + "," + smart_str(Completed_by) + "," + smart_str(Extent_of_Analysis) + "," + smart_str(Population_Frequency) + "," + smart_str(Web_based_Literature_Search) + "," + smart_str(Frameshift_Mutation) + "," + smart_str(In_frame_Deletion_or_Insertion) + ","+  smart_str(Nonsense_Mutation) + "," + smart_str(Evolutionary_Conservation) + "," + smart_str(Severity_of_Amino_Acid_Substitution) + "," + smart_str(Splice_Site_Prediction) + "," + smart_str(Functional_Studies)  + "," + smart_str(mRNA_Analysis) + "," + smart_str(Genotype_Phenotype_Correlation) + "," + smart_str(Co_Segregation) + "," + smart_str(De_Novo_Variant) + "," + smart_str(Phase_established_for_Multiple_Variants_in_cis_or_in_trans) + "," + smart_str(Tumour_Studies) + "," + smart_str(Heteroplasmy_and_Mitochondrial_DNA_changes) + "," + smart_str(final_class) + "," +  "\n"
    with open("/home/bioinfo/anakev/wc/sdgs/csvfiletest333777999moreuvforms.csv", "ar") as csv_file:
      csv_file.write(smart_str(final_record))
  except Exception, e:
    f = open('log.txt', 'ar')
    f.write('An exceptional thing happed - %s' % e + docx + "\n\n\n")
    f.close()
    print 'process error this' + docx + ' is not parsable ' 
  #zipf.close() can add this to prevent mommory leacks 
def smart_str(x):
  if isinstance(x, unicode):
    return unicode(x).encode("utf-8")
  elif isinstance(x, int) or isinstance(x, float):
    return str(x)
  return x
check_file("/home/bioinfo/anakev/wc/UVs_fromblobfiles/more_UVs29092017/")
#process("/home/bioinfo/anakev/wc/UVs_fromblobfiles/UVs_27092017/Pro23Leu v5 UV form Sep 2015.docx") whit little modiffication you can use this and run the code on single files. 