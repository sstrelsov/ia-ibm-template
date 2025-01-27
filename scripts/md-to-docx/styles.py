from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def apply_table_style(doc_path: str, table_style: str, save_as: str) -> None:
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

        # 3) (Optional) set table to 100% width
        tbl_pr = table._element.tblPr
        tbl_w = tbl_pr.find(qn("w:tblW"))
        if tbl_w is None:
            tbl_w = OxmlElement("w:tblW")
            tbl_pr.append(tbl_w)
        tbl_w.set(qn("w:type"), "pct")
        tbl_w.set(qn("w:w"), "5000")

        # 4) Toggle style options so that ONLY Header Row is enabled
        tbl_look = tbl_pr.find(qn("w:tblLook"))
        if tbl_look is None:
            tbl_look = OxmlElement("w:tblLook")
            tbl_pr.append(tbl_look)

        # You can set the w:val attribute and the individual toggles below
        # 'firstRow' -> 1 means highlight first row
        # 'noHBand'  -> 1 means "no horizontal banding" (turns off banded rows)
        # 'noVBand'  -> 1 means "no vertical banding" (turns off banded columns), etc.
        # lastRow=0, firstColumn=0, lastColumn=0 disable those styles.
        tbl_look.set(qn("w:val"), "04A0")  # Not strictly required but often included
        tbl_look.set(qn("w:firstRow"), "1")
        tbl_look.set(qn("w:lastRow"), "0")
        tbl_look.set(qn("w:firstColumn"), "0")
        tbl_look.set(qn("w:lastColumn"), "0")
        tbl_look.set(qn("w:noHBand"), "1")
        tbl_look.set(qn("w:noVBand"), "1")

        print(f"[INFO] Applied style '{table_style}' + header-row-only for table {i+1}")

    doc.save(save_as)
    print(f"[INFO] Updated document saved as '{save_as}'")
