import fitz

doc = fitz.open("runs/DOC0000074469_20261002_1045/build/verify_test.pdf")
print("Page count:", doc.page_count)
for i in range(doc.page_count):
    p = doc[i]
    txt = p.get_text("text").strip()
    first_line = txt.split("\n")[0] if txt else "(empty)"
    print(f"Page {i+1}: size={p.rect.width:.1f}x{p.rect.height:.1f}, text_len={len(txt)}: {first_line[:60]}")
