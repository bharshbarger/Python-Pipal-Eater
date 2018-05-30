#!/usr/bin/env python
"""module to generate a docx report"""

#https://python-docx.readthedocs.io/en/latest/user/text.html
#https://python-docx.readthedocs.io/en/latest/user/quickstart.html
import time
import docx
from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches
#from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn

class Reportgen(object):
    """class to generate a docx report from module output"""
    def __init__(self):
        self.today = time.strftime("%m/%d/%Y")

    def run(self,\
        total,\
        unique,\
        top_10,\
        top_10_base,\
        lengths,\
        counts,\
        one_to_six,\
        one_to_eight,\
        more_than_eight,\
        trailing_digit,\
        trailing_number,\
        last_1digit,\
        last_2digit,\
        last_3digit,\
        last_4digit,\
        last_5digit,\
        charset,\
        charset_ordering):

        

        print('\n[+] Generating Appendix')

        #start doc object
        self.document = docx.Document()

        #add paragraph
        paragraph = self.document.add_paragraph()

        #first paragraph is boilerplate explaining this section
        run_paragraph = paragraph.add_run('\nThis Appendix Contains Analysis of Cracked Passwords.\n')
        font = run_paragraph.font
        font.name = 'Hind'
        font.size = Pt(11)

        #create total passwords and unique table
        first_table = self.document.add_table(rows=0, cols=1)
        
        #try to do autofit, doesnt seem to be working
        #first_table.style = 'Table Grid'
        #first_table.allow_autofit
        
        cells = first_table.add_row().cells
        cells[0].text = str(total)
        cells = first_table.add_row().cells
        cells[0].text = str(unique)
        
        paragraph = self.document.add_paragraph()


        #second table
        second_table = self.document.add_table(rows=0, cols=1)
        
        for i, val in enumerate(top_10):
            cells = second_table.add_row().cells
            cells[0].text = str(top_10[i])
        

        paragraph = self.document.add_paragraph()

        #third table
        third_table = self.document.add_table(rows=0, cols=1)
        for i, val in enumerate(top_10_base):
            cells = third_table.add_row().cells
            cells[0].text = str(top_10_base[i])
    
        paragraph = self.document.add_paragraph()


        #fourth table
        fourth_table = self.document.add_table(rows=0, cols=1)
        for i, val in enumerate(lengths):
            cells = fourth_table.add_row().cells
            cells[0].text = str(lengths[i])
    
        paragraph = self.document.add_paragraph()


        self.document.save('Pipal_Appendix.docx'.format())
        print('[+] Appendix Generated! Yay!')
        print('''   +      o     +              o   
    +             o     +       +
    o          +
        o  +           +        +
        +        o     o       +        o
        -_-_-_-_-_-_-_,------,      o 
        _-_-_-_-_-_-_-|   /\_/\  
        -_-_-_-_-_-_-~|__( ^ .^)  +     +  
        _-_-_-_-_-_-_-""  ""      
        +      o         o   +       o
            +         +
            o        o         o      o     +
                o           +
                +      +     o        o      +''')