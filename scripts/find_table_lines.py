import fitz

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')

def analyze_table_lines(page_num):
    page = doc[page_num]
    drawings = page.get_drawings()
    h_lines = []
    v_lines = []
    for d in drawings:
        r = d["rect"]
        w = r.x1 - r.x0
        h = r.y1 - r.y0
        if w > 20 and h <= 2:
            h_lines.append((round(r.y0, 1), round(r.x0, 1), round(r.x1, 1)))
        elif h > 10 and w <= 2:
            v_lines.append((round(r.x0, 1), round(r.y0, 1), round(r.y1, 1)))
    
    # Sort and group unique positions
    unique_y = sorted(list(set(y for y, x0, x1 in h_lines)))
    unique_x = sorted(list(set(x for x, y0, y1 in v_lines)))
    print(f"=== Page {page_num + 1} ===")
    print(f"H lines Y positions ({len(unique_y)}): {unique_y}")
    print(f"V lines X positions ({len(unique_x)}): {unique_x}")

for p in range(4):
    analyze_table_lines(p)
