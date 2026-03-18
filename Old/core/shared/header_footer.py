# =============================================================================
# core/shared/header_footer.py
# PURPOSE: Inject the exact TMSL header and footer into any generated docx.
#
# This is SHARED across all three feedback types — Course-Exit, Faculty
# End-Term, and Faculty Mid-Term all get the same TMSL header/footer.
#
# HOW IT WORKS:
#   The header and footer in the original TMSL documents use complex
#   Word drawing XML (anchored groups with horizontal lines and right-aligned
#   text) that cannot be recreated via the python-docx API.
#   So we store the exact original XMLs in assets/ and transplant them
#   directly into every generated file via zip manipulation.
#
# USAGE:
#   from core.shared.header_footer import inject_header_footer
#   inject_header_footer("output/MyDoc.docx")   # modifies file in-place
# =============================================================================

import os
import re
import zipfile
import shutil


def inject_header_footer(docx_path: str) -> None:
    """
    Inject the TMSL header and footer into a saved docx file.
    Modifies the file in-place.

    Header: top-right italic —
        "Techno Main Salt Lake"
        "EM-4/1, Sector-V, Salt Lake City, Kolkata- 700091"
        + thick horizontal line below

    Footer: thick horizontal line above —
        "© TMSL" italic left  |  "Page X of Y" italic right

    Args:
        docx_path: full path to the .docx file to modify

    Returns:
        None — file is modified in-place.
        Prints a warning and returns without error if assets are missing.
    """
    # Assets live two levels up from this file: core/shared/ -> core/ -> project/
    assets_dir = os.path.join(os.path.dirname(__file__), '..', '..', 'assets')
    hdr_path   = os.path.join(assets_dir, 'header1.xml')
    ftr_path   = os.path.join(assets_dir, 'footer1.xml')
    img_path   = os.path.join(assets_dir, 'image5.png')

    if not os.path.exists(hdr_path):
        print(f"  ⚠  assets/header1.xml not found — header/footer skipped")
        return

    with open(hdr_path, 'rb') as f: header_xml = f.read()
    with open(ftr_path, 'rb') as f: footer_xml = f.read()
    with open(img_path, 'rb') as f: image5     = f.read()

    # Both header and footer reference image5.png via rId1
    hf_rels = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" '
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        b'Target="media/image5.png"/>'
        b'</Relationships>'
    )

    # Read all files from the docx zip into memory
    with zipfile.ZipFile(docx_path, 'r') as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    # --- Inject header, footer, their rels, and the supporting image ---
    files['word/header1.xml']            = header_xml
    files['word/footer1.xml']            = footer_xml
    files['word/_rels/header1.xml.rels'] = hf_rels
    files['word/_rels/footer1.xml.rels'] = hf_rels
    files['word/media/image5.png']       = image5

    # --- Add header/footer relationships to document.xml.rels ---
    doc_rels    = files['word/_rels/document.xml.rels'].decode('utf-8')
    existing    = re.findall(r'Id="rId(\d+)"', doc_rels)
    max_id      = max(int(x) for x in existing) if existing else 10
    hdr_rid     = f"rId{max_id + 1}"
    ftr_rid     = f"rId{max_id + 2}"
    img_rid     = f"rId{max_id + 3}"

    new_rels = (
        f'<Relationship Id="{hdr_rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        f'Target="header1.xml"/>'
        f'<Relationship Id="{ftr_rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        f'Target="footer1.xml"/>'
        f'<Relationship Id="{img_rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="media/image5.png"/>'
    )
    doc_rels = doc_rels.replace('</Relationships>', new_rels + '</Relationships>')
    files['word/_rels/document.xml.rels'] = doc_rels.encode('utf-8')

    # --- Wire header/footer into sectPr in document.xml ---
    doc_xml = files['word/document.xml'].decode('utf-8')
    # Remove any stale references first
    doc_xml = re.sub(r'<w:headerReference[^/]*/>', '', doc_xml)
    doc_xml = re.sub(r'<w:footerReference[^/]*/>', '', doc_xml)
    hf_refs = (
        f'<w:headerReference r:id="{hdr_rid}" w:type="default"/>'
        f'<w:footerReference r:id="{ftr_rid}" w:type="default"/>'
    )
    # Insert immediately after opening <w:sectPr> or <w:sectPr ...>
    doc_xml = re.sub(r'(<w:sectPr(?:\s[^>]*)?>)', r'\1' + hf_refs, doc_xml, count=1)
    files['word/document.xml'] = doc_xml.encode('utf-8')

    # --- Write back via temp file then replace original ---
    tmp_path = docx_path + '.tmp'
    with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    shutil.move(tmp_path, docx_path)
