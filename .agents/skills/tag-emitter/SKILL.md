---
name: tag-emitter
description: Emits canonical documentation tags ([stamp:], [hw:], [signature], etc.) and unicode symbols per Appendix A & B. Use when building or reviewing DOCX output.
---

# Tag Emitter Skill

## When to use
Any time a Word document (.docx) is being built, extracted, reviewed, or verified for pixel-faithful and semantic fidelity.

---

## Canonical Tags (Appendix A — emit verbatim, no paraphrasing)

Emit exactly as written, including brackets, colons, and spacing.
`xxxxxxxxxxx` is filled with extracted text; if nothing extractable, emit the tag name alone (e.g. `[seal:]`).

| Tag | When to emit | Detailed Usage & Arabic Guidance |
| :--- | :--- | :--- |
| `[emblem:]` | Official emblems (pharmacy, hospital, state, government body). | لوجود شعارات رسمية (مثل صيدلية، مستشفى، شعار دولة أو هيئة حكومية). |
| `[redacted]` | Unclear text hidden by paper, scratch, or deliberate block. | كلام مش واضح أو مستخفي بورقة أو شخطبة أو محجوب عمداً. |
| `[PPD]` | Pre-Printed Document / Data. | بيانات مطبوعة مسبقاً في أصل النموذج أو المستند. |
| `[CCI]` | Chamber of Commerce and Industry (or per document context). | غرفة التجارة والصناعة أو حسب سياق الوثيقة. |
| `[logo: xxxxxxxxxxx]` | Logo containing readable text — write ONLY the text, no image. | لوجو يحتوي على كلام مكتوب، يُكتب الكلام مكان `xxxxxxxxxxx` بدون رسم أو إدراج صورة اللوجو. |
| `[signature]` | Handwritten signature of a person. | توقيع يدوي لشخص داخل المستند أو الصور. |
| `[initials]` | Small initials-style signature / visa. | توقيع مصغر / تأشيرة بالأحرف الأولى (فورمة مصغرة). |
| `[stamp: xxxxxxxxxxx]` | Stamps or postal marks — extract the text inside. | أختام أو طوابع بريد، ويُفرغ النص المكتوب داخل الختم مكان `xxxxxxxxxxx`. |
| `[seal:]` | Seal with a person's name or official personal seal (e.g. notary). | إذا كان الختم يحتوي على اسم الشخص أو الختم الرسمي للشخص (مثل ختم كاتب العدل). |
| `[electronic signature: xxxxxxxxxxx]` | Explicit electronic signature; extract data. | توقيع إلكتروني صريح مع تفريغ بياناته إن وجدت. |
| `[e-signature]` | Generic e-signature; include text if extractable. | إمضاء إلكتروني عام، ولو بداخله نص يمكن استخراجه يُكتب. |
| `[digital signature]` | Technically certified digital signature. | توقيع رقمي معتمد تقنياً وموثق برمجياً. |
| `[illegible] xxxxxxxxxxx` | Fully unclear writing (best guess or blank). | كتابة غير واضحة تماماً بسبب رداءة جودة الصفحة أو التصوير (مع كتابة ما يمكن تخمينه أو تركه فارغاً). |
| `[illegible section]` | Entire paragraph/section unreadable. | مقطع أو فقرة كاملة غير ظاهرة أو غير مقروءة. |
| `[illegible line]` | Entire line unreadable. | سطر كامل غير مقروء أو غير ظاهر. |
| `[watermark: xxxxxxxxxx]` | Watermark — usually in Header with its readable text. | علامة مائية؛ تُكتب عادة في الترويسة (Header) مع تفريغ النص المقروء منها. |
| `[icon]` | Any small icon or graphic symbol. | أي أيقونة أو رمز جرافيكي صغير. |
| `[cut off text] xxxxxx` | Text cut at page edge — visible characters only. | كلام مقطوع عند حافة الورقة؛ نكتب الحروف الظاهرة فقط مكان `xxxxxx`. |
| `[hw: xxxxxx]` | Handwritten text — render content in *Italic*. | كلام مكتوب بخط اليد وليس رقمياً (digital)؛ ويوضع النص بتنسيق مائل (*Italic*). |
| `[barcode: xxxx]` | Barcode — do NOT draw; extract numeric/alpha content. | باركود؛ لا يتم رسمه، وإنما تُستخرج الأرقام أو الحروف التابعة له وتُكتب بجانبه. |
| `[QR code]` | QR code. | رمز الاستجابة السريعة QR code. |
| `[blank page in source]` | Page is blank in original PDF. | صفحة فارغة في ملف الـ PDF الأصلي. |
| `undefined` | Empty boxes / fields left blank. | عندما توجد مربعات أو حقول فارغة يُطلب ملؤها وتُركت فارغة. |
| `[handwritten text is indicated in italics]` | If file has MANY handwritten parts — emit ONCE in Header only. | إذا كان الملف يحتوي على الكثير من النصوص المكتوبة بخط اليد، توضع هذه العبارة في الـ Header فقط. |

---

## Canonical Symbols (Appendix B — Unicode only)

### Checkboxes & Bullets
- **Filled square:** `■` (`U+25A0` / Alt + 9632)
- **Empty square:** `□` (`U+25A1` / Alt + 9633)
- **Filled circle:** `●` (`U+25CF` / Alt + 9679)
- **Empty circle:** `○` (`U+25CB` / Alt + 9675)
- **Multiplication / X mark:** `(×)` or `×` (`U+00D7`)

### Punctuation & Special Characters
- **Vertical bar:** `|`
- **According to (German):** `gemäß`
- **Paragraph / Legal article:** `§` (`U+00A7`)
- **Numero sign:** `№` (`U+2116`)

### Time Zones (equivalent within each group)
- **US Eastern:** `ET` / `EDT` / `EST`
- **Europe/UK:** `CET` / `BST` / `GMT`

### Medical & Scientific Notation
- **Pharmacovigilance:** `pharmacovigilance`
- **Mean ± SD:** `(x̄±s)` (combining macron `\u0304` on `x`)
- **Greek letters:**
  - Alpha: `α` (`U+03B1`)
  - Beta: `β` (`U+03B2`)
  - Gamma: `γ` (`U+03B3`)
- **Extended Arabic Hah (Tatweel/Kashida variant):** `هـ`
- **Common fractions:**
  - One quarter: `¼` (`U+00BC`)
  - One half: `½` (`U+00BD`)
  - One third: `⅓` (`U+2153`)
  - Three quarters: `¾` (`U+00BE`)

---

## Strict Grammar & Syntax Rules
1. **Always square brackets:** Always enclose tags in `[ ]`. Never use `<stamp>`, `(signature)`, or `{emblem}`.
2. **Payload format:** Parameterized tags must follow `[tag: payload]` (colon followed by exactly one space).
3. **Never translate tag names:** Tag identifiers (`emblem`, `signature`, `stamp`, `hw`, etc.) must remain in English as specified.
4. **No line-splitting:** Never break a tag across a newline boundary.
5. **No nested tags:** Tags must not be nested (e.g. `[stamp: [signature]]` is strictly invalid; emit sequentially or extract text).
6. **Handwriting styling:** Any text within `[hw: ...]` MUST be formatted with *Italic* font in Word.
7. **No ASCII substitutes:** Never substitute `x` for `×`, `B` for `β`, or `-` for `ـ`. Always preserve combining characters like `x̄`.
