# -*- coding: utf-8 -*-

import io
import re
import pdfplumber

from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from ExcelHandler import ExcelList
from ExcelHandler import TransformExcelListToDict
from tqdm import tqdm

def LoadExcelListFromPDF(pdf_path):
    data_list = list()
    with pdfplumber.open(pdf_path) as pdf:
        with tqdm(total = len(pdf.pages)) as pbar:
            pbar.set_description('\tHave completed：')
            for (i, page) in enumerate(pdf.pages):
                for (j, annot) in enumerate(page.annots):
                    ans_data = ExcelList()
                    ans_data.text = annot.get('contents')
                    if(ans_data.text == None):
                        continue
                    ans_data.page = annot.get('page_number')
                    assert i == int(ans_data.page) -1
                    ans_data.x_position = annot.get('x0')
                    ans_data.y_position = annot.get('y0')
                    ans_data.height = annot.get('height')
                    ans_data.width = annot.get('width')
                    ans_data.color = annot.get('data').get('C')
                    font = annot.get('data').get('DS')
                    fontinfo = str(font).split(';')
                    fontsize = float(re.findall("\d+\.\d+|\d+", fontinfo[0])[0])
                    fontcolor = str(annot.get('data').get('RC'))
                    fontcolor = str(re.findall("#\w\w\w\w\w\w", fontcolor)[-1])
                    ans_data.fontsize = fontsize
                    ans_data.fontcolor = fontcolor
                    data_list.append(ans_data)
                pbar.update(1)
    return data_list

def _TransformText(text, max_count):
    original_text_list = text.strip().split('\r')
    text_list = list()
    for string in original_text_list:
        if(len(string) < max_count):
            text_list.append(string)
        else:
            for i in range(len(string) // max_count):
                text_list.append(string[i * max_count : (i + 1) * max_count])
            if(len(string) % max_count != 0):
                text_list.append(string[(i + 1) * max_count : ])
    return "\n".join(text_list)

def WriteExcelListToPDF(data_list, original_pdf_path, saved_pdf_path):
    data_dict = TransformExcelListToDict(data_list)
    pdfmetrics.registerFont(TTFont("SimSun", "simsunb.ttf"))
    ParagraphStyle.defaults['wordWrap'] = 'CJK'
    output = PdfFileWriter()
    original_packet_list = list()
    packet_list = list()
    with open(original_pdf_path, "rb") as f:
        source = PdfFileReader(f, "rb")
        with tqdm(total = source.getNumPages()) as pbar:
            pbar.set_description('\tHave completed：')
            for i in range(source.getNumPages()):
                original_packet = io.BytesIO()
                original_packet.seek(0)
                page = source.getPage(i)
                if('/Annots' in page.keys()):
                    page.pop('/Annots')
                original_packet_list.append(original_packet)
                
                if(i not in list(data_dict.keys())):
                    output.addPage(page)
                else:
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=A4)
                    can.setFont("SimSun", 18)
                    for item in data_dict[i]:
                        x, y = item['x_position'], item['y_position']
                        height, width = item['height'], item['width']
                        r, g, b = item['color']
                        fontsize, fontcolor = item['fontsize'], item['fontcolor']
                        
                        max_count = int(width / (fontsize * (1 / 2.1)))
                        text = _TransformText(item['text'], max_count)
                        
                        can.setFillColorRGB(r=r, g=g, b=b)
                        can.rect(x, y, width, height, stroke = 1, fill = 1)
                        can.setFillColorRGB(r=0, g=0, b=0)
                        t = can.beginText()
                        t.setFont("SimSun", fontsize)
                        t.setFillColor(fontcolor)
                        t.setTextOrigin(x, y + fontsize * (4 / 3) *(len(text.split('\n')) - 1) + 5)
                        t.textLines(text)
                        can.drawText(t)
                    can.save()
                    
                    packet.seek(0)
                    new_pdf = PdfFileReader(packet)
                    page.mergePage(new_pdf.getPage(0))
                    output.addPage(page)
                    packet_list.append(packet)
                pbar.update(1)
        with open(saved_pdf_path, "wb") as outputStream:
            output.write(outputStream)
    for packet in packet_list:
        packet.close()
    for original_packet in original_packet_list:
        original_packet.close()