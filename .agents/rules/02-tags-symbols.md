---
description: Contractual tag lexicon (Appendix A) and canonical symbol set (Appendix B) for transcription and document building
trigger: model_decision
---

# APPENDIX A — TAG LEXICON (canonical, do not paraphrase)

Emit exactly as written, including brackets, colons, spacing.
`xxxxxxxxxxx` is filled with extracted text; if nothing extractable,
emit the tag name alone (e.g. `[seal:]`).

| Tag | When to emit | Description / Arabic Notes |
| :--- | :--- | :--- |
| `[emblem:]` | Official emblems (pharmacy, hospital, state, government body). | لوجود شعارات رسمية (مثل صيدلية، مستشفى، شعار دولة أو هيئة حكومية). |
| `[redacted]` | Unclear text hidden by paper, scratch, or deliberate block. | كلام مش واضح أو مستخفي بورقة أو شخطبة أو محجوب عمداً. |
| `[PPD]` | Pre-Printed Document / Data. | بيانات مطبوعة مسبقاً في النموذج الأصلي. |
| `[CCI]` | Chamber of Commerce and Industry (or per document context). | غرفة التجارة والصناعة أو حسب سياق الوثيقة. |
| `[logo: xxxxxxxxxxx]` | Logo containing readable text — write ONLY the text, no image. | لوجو يحتوي على نص مقروء؛ يُكتب النص فقط دون إدراج صورة. |
| `[signature]` | Handwritten signature of a person. | توقيع يدوي لشخص داخل المستند أو الصور. |
| `[initials]` | Small initials-style signature / visa. | توقيع مصغر / تأشيرة بالأحرف الأولى (فورمة مصغرة). |
| `[stamp: xxxxxxxxxxx]` | Stamps or postal marks — extract the text inside. | أختام أو طوابع بريد؛ يُفرغ النص المقروء داخل الختم. |
| `[seal:]` | Seal with a person's name or official personal seal (e.g. notary). | ختم رسمي لشخص أو جهة رسمية (مثل ختم كاتب العدل). |
| `[electronic signature: xxxxxxxxxxx]` | Explicit electronic signature; extract data. | توقيع إلكتروني صريح مع تفريغ بياناته المقروءة. |
| `[e-signature]` | Generic e-signature; include text if extractable. | إمضاء إلكتروني عام؛ ولو بداخله نص يُكتب. |
| `[digital signature]` | Technically certified digital signature. | توقيع رقمي معتمد تقنياً وموثق برمجياً. |
| `[illegible] xxxxxxxxxxx` | Fully unclear writing (best guess or blank). | كتابة غير واضحة تماماً؛ يكتب التخمين الأقرب أو يُترك فارغاً. |
| `[illegible section]` | Entire paragraph/section unreadable. | مقطع أو فقرة كاملة غير ظاهرة أو غير مقروءة. |
| `[illegible line]` | Entire line unreadable. | سسطر كامل غير مقروء أو غير ظاهر. |
| `[watermark: xxxxxxxxxx]` | Watermark — usually in Header with its readable text. | علامة مائية؛ تُكتب عادة في الترويسة مع النص المقروء منها. |
| `[icon]` | Any small icon or graphic symbol. | أي أيقونة أو رمز جرافيكي صغير. |
| `[cut off text] xxxxxx` | Text cut at page edge — visible characters only. | كلام مقطوع عند حافة الورقة؛ نكتب الحروف الظاهرة فقط. |
| `[hw: xxxxxx]` | Handwritten text — render content in *Italic*. | كلام مكتوب بخط اليد؛ ويوضع النص بتنسيق مائل (*Italic*). |
| `[barcode: xxxx]` | Barcode — do NOT draw; extract numeric/alpha content. | باركود؛ لا يُرسم وإنما تُفرغ أرقامه وحروفه. |
| `[QR code]` | QR code. | رمز الاستجابة السريعة QR code. |
| `[blank page in source]` | Page is blank in original PDF. | صفحة فارغة في ملف الـ PDF الأصلي. |
| `undefined` | Empty boxes / fields left blank. | مربعات أو حقول فارغة تُركت بدون ملء. |
| `[handwritten text is indicated in italics]` | If file has MANY handwritten parts — emit ONCE in Header only. | إذا كان الملف يحتوي على الكثير من النصوص اليدوية، توضع في Header فقط. |

### Grammar Rules
- Always use square brackets `[ ]`.
- Payload uses `: ` (colon + space) — e.g. `[stamp: Approved]`.
- Never translate tag names to other languages.
- Never split tags across lines.
- Never nest tags (e.g. `[stamp: [signature]]` is strictly forbidden).
- **NEVER emit `[logo:]` without readable text.** If a logo, emblem, or mobile badge (e.g., App Store, Google Play) contains no readable text, emit `[icon]` or embed the graphic directly. Never place isolated tag placeholders inside arbitrary border boxes.


---

# APPENDIX B — SYMBOL SET (canonical characters)

## Checkboxes & Bullets
- Filled square `■` (U+25A0 / Alt + 9632)
- Empty square `□` (U+25A1 / Alt + 9633)
- Filled circle `●` (U+25CF / Alt + 9679)
- Empty circle `○` (U+25CB / Alt + 9675)
- Multiplication / X mark: `(×)` or `×`

## Punctuation & Special
- Vertical bar `|`
- "According to" (German): `gemäß`
- Paragraph / legal article: `§`
- Numero sign: `№`

## Time Zones (equivalent within each group)
- US Eastern: `ET` / `EDT` / `EST`
- Europe/UK: `CET` / `BST` / `GMT`

## Medical & Scientific
- Pharmacovigilance: `pharmacovigilance`
- Mean ± SD: `(x̄±s)`
- Greek letters: `α` `β` `γ`
- Extended Arabic Hah: `هـ`
- Common fractions: `¼` `½` `⅓` `¾`

## Rendering Constraints
- Prefer Unicode characters over ASCII approximations.
- Never substitute `x` for `×`, `B` for `β`, or `-` for `ـ`.
- Preserve the combining macron on `x̄`.
- If a font lacks a glyph, fall back per config — do not silently replace.
