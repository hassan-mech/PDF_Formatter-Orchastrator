import fitz

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')

for pno in range(9, len(doc)):
    page = doc[pno]
    infos = page.get_image_info()
    min_x = min(info['bbox'][0] for info in infos)
    min_y = min(info['bbox'][1] for info in infos)
    max_x = max(info['bbox'][2] for info in infos)
    max_y = max(info['bbox'][3] for info in infos)
    print(f"Page {pno+1}: count={len(infos)}, bbox=({min_x:.1f}, {min_y:.1f}) -> ({max_x:.1f}, {max_y:.1f}), size={max_x-min_x:.1f}x{max_y-min_y:.1f}")
