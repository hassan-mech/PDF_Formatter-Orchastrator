# Compare and inspect exact text for pages 4 to 13

for p in range(4, 14):
    with open(f'pdf_pages_judgment/page_{p:02d}_ocr.txt', 'r', encoding='utf-8') as f:
        ocr_lines = [l.strip() for l in f.readlines() if l.strip()]
    print(f"\n==================== PAGE {p} ====================")
    for l in ocr_lines:
        print("  ", l)
