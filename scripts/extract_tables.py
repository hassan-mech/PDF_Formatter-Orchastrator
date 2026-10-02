import fitz

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')

# Let's inspect pages 1 to 4 tables in detail
for pno in range(4):
    page = doc[pno]
    tabs = page.find_tables()
    print(f"=== Page {pno+1} Tables: {len(tabs.tables)} ===")
    for t_idx, tab in enumerate(tabs.tables):
        df = tab.extract()
        print(f"--- Table {t_idx+1} ({len(df)} rows, {len(df[0])} cols) ---")
        for row_idx, row in enumerate(df):
            row_clean = [cell.strip().replace('\n', ' ') if cell else '' for cell in row]
            print(f"  Row {row_idx}: {row_clean}")
