
from docx import Document
from docx.shared import Pt
import re

# doc = Document('test_file.docx')
# style_f = doc.styles['Footer']
# font_f = style_f.font
# font_f.name = 'Arial'
# font_f.bold = False
# font_f.size = Pt(10)

# footer = doc.sections[0].footer
# paragraph_f = footer.paragraphs[0]
# paragraph_f.text = "FRO-018 (NOV 08, 2022)  	© Kings’s Printer for Ontario, 2022" # insert new value here.
# paragraph_f.style = doc.styles['Footer']# this is what changes the style

# doc.save('test_file.docx')
# On the pdf format, I tested three approaches:





# 1) using API
# 2) converting.txt files from PDF
# 3) Making use of the pdfpy library





# I can confirm that none of these deliver the optimal results.





# --------------------------------------------------------------------





# With regard to the docx file, I used the docx library framework in Python, and I think it looks good.





# 1) It can create a new footer for the file, but it cannot search for 'Queen'.
# 2) In the paragraph(body) part of the file, it will replace the Queen word with 'King',
# and hence replacement works





# ---------------------------------------------------------------------------





Additionally, I am also exploring another approach that makes use of the C#.NET Framework.




def Text_Replacer(document, word , replace):

    for p in document.paragraphs:
        if word.search(p.text):
            inline = p.runs

            for i in range(len(inline)):
                if word.search(inline[i].text):
                    text = word.sub(replace, inline[i].text)
                    inline[i].text = text
    doc = document             
    style_f = doc.styles['Footer']
    font_f = style_f.font
    font_f.name = 'Arial'
    font_f.bold = False
    font_f.size = Pt(10)
    footer = doc.sections[0].footer
    paragraph_f = footer.paragraphs[0]
    paragraph_f.text = "FRO-018 (NOV 08, 2022)  	© Kings’s Printer for Ontario, 2022" # insert new value here.
    paragraph_f.style = doc.styles['Footer']# this is what changes the style





words = re.compile(r"Queen")
replace1 = r"King"
filename = "test_file.docx"
doc = Document(filename)
Text_Replacer(doc, words, replace1)
doc.save('test_file.docx')
