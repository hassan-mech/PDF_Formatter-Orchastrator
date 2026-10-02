import fitz

orig = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')
exp = fitz.open('output_word_exported.pdf')
print(f"Orig pages: {len(orig)}, Exported pages: {len(exp)}")

for i in range(max(len(orig), len(exp))):
    p_orig = orig[i] if i < len(orig) else None
    p_exp = exp[i] if i < len(exp) else None
    
    t_orig = p_orig.get_text().strip()[:40].replace('\n', ' ') if p_orig else '[NONE]'
    t_exp = p_exp.get_text().strip()[:40].replace('\n', ' ') if p_exp else '[NONE]'
    
    imgs_orig = len(p_orig.get_images()) if p_orig else 0
    imgs_exp = len(p_exp.get_images()) if p_exp else 0
    
    print(f"P{i+1}: Orig(t='{t_orig[:25]}', img={imgs_orig}) | Exp(t='{t_exp[:25]}', img={imgs_exp})")
