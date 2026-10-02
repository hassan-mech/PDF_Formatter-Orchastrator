import os
import zipfile
import hashlib
import xml.etree.ElementTree as ET

docx_path = r"output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx"
run_docx_path = r"runs/DOC0000074469_20261002_150418/build/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx"

print(f"=== CHECK 1: STATIC ANALYSIS & OPENXML AUTHENTICITY ===")

# Hash comparison
with open(docx_path, "rb") as f:
    h_out = hashlib.sha256(f.read()).hexdigest()
with open(run_docx_path, "rb") as f:
    h_run = hashlib.sha256(f.read()).hexdigest()

print(f"output docx sha256: {h_out}")
print(f"run build docx sha256: {h_run}")
print(f"Match: {h_out == h_run}")

# Zip structure inspection
is_zip = zipfile.is_zipfile(docx_path)
print(f"Is valid ZIP archive: {is_zip}")

with zipfile.ZipFile(docx_path, "r") as zf:
    namelist = zf.namelist()
    print(f"Total entries in ZIP: {len(namelist)}")
    required_parts = [
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml",
        "word/_rels/document.xml.rels",
        "word/styles.xml",
        "word/settings.xml"
    ]
    for part in required_parts:
        print(f"  Presence of {part}: {part in namelist}")
    
    # Check media files
    media_files = [n for n in namelist if n.startswith("word/media/")]
    print(f"Media files count: {len(media_files)}")
    for mf in sorted(media_files):
        info = zf.getinfo(mf)
        print(f"  {mf}: size={info.file_size} bytes")

    # Inspect document.xml
    doc_xml_bytes = zf.read("word/document.xml")
    print(f"word/document.xml uncompressed size: {len(doc_xml_bytes)} bytes")
    root = ET.fromstring(doc_xml_bytes)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    body = root.find("w:body", ns)
    if body is not None:
        p_count = len(body.findall("w:p", ns))
        tbl_count = len(body.findall("w:tbl", ns))
        sect_count = len(root.findall(".//w:sectPr", ns))
        print(f"XML Elements: Paragraphs={p_count}, Tables={tbl_count}, sectPr={sect_count}")
