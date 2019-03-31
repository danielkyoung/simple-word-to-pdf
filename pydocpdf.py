# -*- coding: utf-8 -*-
"""
Created on Tue Feb 26 19:49:58 2019

@author: danie
"""
from fpdf import FPDF
from docx import Document
def ttf_conversion (name,style):
    keydict = {'Times New Roman':'times',\
               'Arial':'Arial',\
               'Calibri':'calibri',\
               'Coureir New':'courier'}
    styledict = {'B':'b',
                 'BU':'b',
                 'BI':'z',
                 'BIU':'z',
                 'IU':'i',
                 'I':'i'}
    if name == 'Calibri':
        return keydict[name]+styledict[style]
    else:
        try:
            return keydict[name]
        except:
            return 'Arial'
def convert_docx(docx_directory,pdf_directory):
    document = Document(docx_directory)
    pdf = FPDF('P','pt','A4')
    pdf.add_page()
    pdf.set_margins(document.sections[0].left_margin.pt,document.sections[0].top_margin.pt,document.sections[0].right_margin.pt)
    paragraphs = document.paragraphs
    for paragraph in paragraphs:
        run_list = []
        if paragraph.paragraph_format.line_spacing is None:
            line_spacing = 1.15
        else:
            line_spacing = paragraph.paragraph_format.line_spacing
        for run in paragraph.runs:
            run_dict = {}
            run_dict['text'] = run.text.replace(u"\u2019",u"\u0027")\
                .replace(u"\u201c",u"\u0022")\
                .replace(u"\u201d",u"\u0022")\
                .replace(u"\u0009",u"\u0020"*4)
            if run.font.name is None:
                run_dict['font']='calibri'
            else:
                run_dict['font']=run.font.name
            if run.font.size is None:
                run_dict['size']=11.0
            else:
                run_dict['size']=run.font.size.pt
            if run.font.bold is None:
                run_dict['style']=''
            else:
                run_dict['style']='B'
            if run.font.italic is None:
                pass
            else:
                run_dict['style']+='I'
            if run.font.underline is None:
                pass
            else:
                run_dict['style']+='U'
            run_list.append(run_dict)
        if paragraph.paragraph_format.alignment is None or paragraph.paragraph_format.alignment == 'LEFT':
            for item in run_list:
                try:
                    pdf.set_font(ttf_conversion(item['font'],item['style']),item['style'],item['size'])
                except:
                    directory = "C:\\Windows\\Fonts\\"+ ttf_conversion(item['font'],item['style'])+'.ttf'
                    pdf.add_font(item['font'],item['style'],directory, uni=True)
                    pdf.set_font(item['font'],item['style'],item['size'])
                pdf.write(item['size']*line_spacing,item['text'])
        else:
            run = run_list[0]
            pdf.set_font(ttf_conversion(run['font'],run['style']),run['style'],run['size'])
            text = ''
            for item in run_list:
                text += item['text']
            pdf.multi_cell(0,run['size']*line_spacing,text,0,'C',False)
        pdf.ln()
        pdf.set_font('Arial','',11)
    pdf.output(pdf_directory)
if __name__ == "__main__":
    docx_directory = input('Enter the document directory\n')
    pdf_directory = input('Enter the pdf directory\n')
    try:
        convert_docx(docx_directory,pdf_directory)
    except:
        print('invalid directory')