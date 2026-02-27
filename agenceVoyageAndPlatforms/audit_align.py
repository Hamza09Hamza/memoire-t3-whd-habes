from docx import Document
from docx.oxml.ns import qn

doc = Document('main_word_styled.docx')

# Check all heading paragraphs and their actual XML jc values
print("=== Heading alignment audit ===\n")

for i, p in enumerate(doc.paragraphs):
    style = p.style.name
    if 'Heading' not in style and i > 24:
        continue
    if i > 24 and 'Heading' not in style:
        continue
    
    text = p.text.strip()
    if not text:
        continue
    
    ppr = p._element.find(qn('w:pPr'))
    
    # Get jc value
    jc = None
    if ppr is not None:
        jc_el = ppr.find(qn('w:jc'))
        jc = jc_el.get(qn('w:val')) if jc_el is not None else None
    
    # Get bidi
    bidi = ppr.find(qn('w:bidi')) if ppr is not None else None
    has_bidi = bidi is not None
    
    # Python-docx alignment
    py_align = p.paragraph_format.alignment
    
    if i <= 24 or 'Heading' in style:
        print(f"[{i:4d}] {style:12s} | jc={str(jc):8s} | bidi={has_bidi} | py_align={py_align} | {text[:50]}")

# Also check style-level settings
print("\n=== Style-level alignment ===")
for sname in ['Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4']:
    if sname in doc.styles:
        s = doc.styles[sname]
        ppr = s.element.find(qn('w:pPr'))
        jc = None
        bidi_s = False
        if ppr is not None:
            jc_el = ppr.find(qn('w:jc'))
            jc = jc_el.get(qn('w:val')) if jc_el is not None else None
            bidi_s = ppr.find(qn('w:bidi')) is not None
        print(f"  {sname:12s} | style_jc={str(jc):8s} | style_bidi={bidi_s}")
