from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def insert_paragraph_after_table(table, text=""):
    """
    Inserts a new paragraph *immediately* after the given table.
    Returns a python-docx Paragraph object that you can then modify
    (e.g. set text, style, etc.).
    """
    # Create a new <w:p> element
    new_p = OxmlElement("w:p")

    # Insert it right after the table's XML
    table._element.addnext(new_p)

    # Wrap the <w:p> element in a python-docx Paragraph object
    paragraph = Paragraph(new_p, table._parent)
    paragraph.text = text
    return paragraph


def apply_table_style(doc_path: str, table_style: str, save_as: str) -> None:
    """
    Open a Word doc, set each table to `table_style`, remove fixed widths,
    set table width to 100%, and enable only header-row styling. Then save.
    """
    doc = Document(doc_path)

    for i, table in enumerate(doc.tables):
        # 1) Apply style
        table.style = table_style
        table.autofit = True

        # 2) Remove any fixed cell width from each cell
        for row in table.rows:
            for cell in row.cells:
                tcPr = cell._tc.get_or_add_tcPr()
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is not None:
                    tcPr.remove(tcW)

        # 3) Set table to 100% width
        tbl_pr = table._element.tblPr
        tbl_w = tbl_pr.find(qn("w:tblW"))
        if tbl_w is None:
            tbl_w = OxmlElement("w:tblW")
            tbl_pr.append(tbl_w)
        tbl_w.set(qn("w:type"), "pct")
        tbl_w.set(qn("w:w"), "5000")  # 50% of page width (approx)

        # 4) Toggle style options so that ONLY Header Row is enabled
        tbl_look = tbl_pr.find(qn("w:tblLook"))
        if tbl_look is None:
            tbl_look = OxmlElement("w:tblLook")
            tbl_pr.append(tbl_look)

        tbl_look.set(qn("w:val"), "04A0")  # Not strictly required, but common
        tbl_look.set(qn("w:firstRow"), "1")
        tbl_look.set(qn("w:lastRow"), "0")
        tbl_look.set(qn("w:firstColumn"), "0")
        tbl_look.set(qn("w:lastColumn"), "0")
        tbl_look.set(qn("w:noHBand"), "1")
        tbl_look.set(qn("w:noVBand"), "1")
        
        insert_paragraph_after_table(table)

        print(f"[INFO] Applied style '{table_style}' to table {i+1} with header-row only")

    doc.save(save_as)
    print(f"[INFO] Updated document saved as '{save_as}'")
    print(f"[INFO] Updated document saved as '{save_as}'")
