# -*- coding: utf-8 -*-

from ExcelHandler import WriteExcelListToExcel
from ExcelHandler import LoadExcelListFromExcel

from PDFHandler import LoadExcelListFromPDF
from PDFHandler import WriteExcelListToPDF

import argparse
import os

parser = argparse.ArgumentParser()
parser.add_argument("--pdf_path", type = str, default = "pdfs")

option = parser.parse_args()

if(option.pdf_path.split('.')[-1] == "pdf"):
    original_pdf_path = option.pdf_path
    short_name = original_pdf_path.split(".")[0]
    saved_pdf_path = "saved_{}.pdf".format(short_name)
    excel_path = "saved_{}.xls".format(short_name)
    
    print("Step 1: Load information from PDF file into list.")
    data_list_pdf = LoadExcelListFromPDF(original_pdf_path)
    print("Step 2: Store list information into xls file.")
    WriteExcelListToExcel(data_list_pdf, excel_path)
    
    print("Step 3: Load information from Excel file into list.")
    data_list_excel = LoadExcelListFromExcel(excel_path)
    print("Step 4: Store list information into original PDF file.")
    WriteExcelListToPDF(data_list_excel, original_pdf_path, saved_pdf_path)
else:
    if(not os.path.exists("saved_pdfs")):
        os.mkdir("saved_pdfs")
    if(not os.path.exists("excels")):
        os.mkdir("excels")
        
    names = [name for name in os.listdir(option.pdf_path) if name.split(".")[-1] == "pdf"]
    for (i, name) in enumerate(names):
        print("Operating index {}, file name is {}".format(i, name))
        original_pdf_path = os.path.join(option.pdf_path, name)
        short_name = os.path.basename(original_pdf_path).split(".")[0]
        saved_pdf_path = os.path.join("saved_pdfs", "saved_{}.pdf".format(short_name))
        excel_path = os.path.join("excels", "saved_{}.xls".format(short_name))
        print("Step 1: Load information from PDF file into list.")
        data_list_pdf = LoadExcelListFromPDF(original_pdf_path)
        print("Step 2: Store list information into xls file.")
        WriteExcelListToExcel(data_list_pdf, excel_path)
        
        print("Step 3: Load information from Excel file into list.")
        data_list_excel = LoadExcelListFromExcel(excel_path)
        print("Step 4: Store list information into original PDF file.")
        WriteExcelListToPDF(data_list_excel, original_pdf_path, saved_pdf_path)