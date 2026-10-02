import fitz

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')

for pno in range(9):
    page = doc[pno]
    rect = page.rect
    # Find drawing paths / table lines
    drawings = page.get_drawings()
    print(f"Page {pno+1}: width={rect.width}, height={rect.height}, drawings={len(drawings)}")
    # get min/max coordinates of drawings and text
    min_x, min_y, max_x, max_y = 1000, 1000, 0, 0
    for d in drawings:
        r = d["rect"]
        min_x = min(min_x, r.x0)
        min_y = min(min_y, r.y0)
        max_x = max(max_x, r.x1)
        max_y = max(max_y, r.y1)
    print(f"  Drawings bbox: ({min_x:.1f}, {min_y:.1f}) -> ({max_x:.1f}, {max_y:.1f})")
    
    t_min_x, t_min_y, t_max_x, t_max_y = 1000, 1000, 0, 0
    for b in page.get_text("blocks"):
        if b[4].strip():
            t_min_x = min(t_min_x, b[0])
            t_min_y = min(t_min_y, b[1])
            t_max_x = max(t_max_x, b[2])
            t_max_y = max(t_max_y, b[3])
    print(f"  Text bbox: ({t_min_x:.1f}, {t_min_y:.1f}) -> ({t_max_x:.1f}, {t_max_y:.1f})")
