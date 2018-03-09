from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt 

name = "Duc Nguyen"
address = "Victoria BC"
phone = "250-884-6325"
mail = "dukeng@uvic.ca"
website="https://dukeng.github.io/"
github="https://github.com/dukeng/"

document = Document()

document.add_heading(name, 2)




contact_info = document.add_paragraph()
# contact_info.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

address_obj = contact_info.add_run(address + " , ")
address_obj.font.size = Pt(24)
contact_info.add_run(phone + " , ")
contact_info.add_run(mail)
contact_info.add_run("\n" + "Website:" + website)
contact_info.add_run("\n" + "Github:" + github)



# p.add_run('bold').bold = True
# p.add_run(' and some ')
# p.add_run('italic.').italic = True

# document.add_heading('Heading, level 1', level=1)
# document.add_paragraph('Intense quote', style='IntenseQuote')

# document.add_paragraph(
#     'first item in unordered list', style='ListBullet'
# )
# document.add_paragraph(
#     'first item in ordered list', style='ListNumber'
# )


# table = document.add_table(rows=1, cols=3)
# hdr_cells = table.rows[0].cells
# hdr_cells[0].text = 'Qty'
# hdr_cells[1].text = 'Id'
# hdr_cells[2].text = 'Desc'

document.add_page_break()

document.save('news.docx')