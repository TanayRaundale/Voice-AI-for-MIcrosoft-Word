from docx import Document
from docx.shared import Pt


def set_margins(doc, top, right, bottom, left):
    sections = doc.sections

    for section in sections:
        section.top_margin = Pt(top)
        section.right_margin = Pt(right)
        section.bottom_margin = Pt(bottom)
        section.left_margin = Pt(left)

def main():

    doc = Document("cloud.docx")   # Create a new Word document

    # Set margins (in points)
    top_margin = 2.54  # 1/2 inch
    right_margin = 1.5
    bottom_margin = 2.86
    left_margin = 3.81

    set_margins(doc, top_margin, right_margin, bottom_margin, left_margin)

    # doc.add_heading('Document generated with Default Margins CM', level=1)
    # doc.add_paragraph("")
    # doc.add_paragraph('This is an Simple Default Margin Example and generates Document...')
    doc.save("Applied_margin.docx")

if __name__ == "__main__":
    main()
