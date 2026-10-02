import pymupdf

doc = pymupdf.open('finished/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf')

for pno in range(1, 6):
    page = doc[pno]
    print(f"\n================ PAGE {pno + 1} ================")
    blocks = [b for b in page.get_text('blocks') if len(b[4].strip()) > 5]
    for b in blocks[:6]:
        print(f"  Block ({round(b[0],1)}, {round(b[1],1)}, {round(b[2],1)}, {round(b[3],1)}): {repr(b[4][:40])}")
