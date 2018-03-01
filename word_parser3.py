from docx import Document
from docx.shared import Inches


def main():
    document_read = raw_input("Enter file name or path\n")
    document = Document(document_read)
    num = 1
    for para in document.paragraphs:
        name = "Topic"+str(num -1)+".docx"
        if para.style.name == 'Heading 1' and num == 1:
            doc = Document()
            sty = int(para.style.name[8])
            doc.add_heading(para.text, sty)
            num = num + 1
        elif para.style.name == 'Heading 1' and num > 1:
            doc.save(name)
            doc = Document()
            doc.add_heading(para.text, 1)
            num = num + 1
        elif para.style.name != 'Heading 1' and para.style.name.startswith('Heading'):
            sty = int(para.style.name[8])
            doc.add_heading(para.text, sty)
        else:
            doc.add_paragraph(para.text)
    doc.save(name)
main()
