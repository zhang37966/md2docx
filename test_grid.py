from docx import Document
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml
from docx.enum.style import WD_STYLE_TYPE

def test():
    doc = Document()
    for style in doc.styles:
        if style.type == WD_STYLE_TYPE.PARAGRAPH:
            pPr = style._element.get_or_add_pPr()
            
            # Remove existing if any
            for tag in ('w:snapToGrid', 'w:adjustRightInd'):
                existing = pPr.find(qn(tag))
                if existing is not None:
                    pPr.remove(existing)
                    
            pPr.append(parse_xml(f'<w:snapToGrid {nsdecls("w")} w:val="0"/>'))
            pPr.append(parse_xml(f'<w:adjustRightInd {nsdecls("w")} w:val="0"/>'))

    p = doc.add_paragraph('Test Paragraph')
    doc.save('test_grid.docx')
    print('Success')

if __name__ == '__main__':
    test()
