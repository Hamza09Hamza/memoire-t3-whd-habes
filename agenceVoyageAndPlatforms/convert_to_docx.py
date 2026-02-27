#!/usr/bin/env python3
"""
Convert Arabic LaTeX thesis to DOCX with proper RTL and Arabic font support.
Steps:
1. Consolidate all .tex files into one cleaned LaTeX file
2. Create a reference DOCX template with RTL/Arabic settings
3. Run pandoc to convert
"""

import re
import os
import subprocess

BASE = os.path.dirname(os.path.abspath(__file__))

# ─── Citation key → readable Arabic text mapping ──────────────────────────
CITATIONS = {
    'buhalis2020': 'بوهاليس، 2020',
    'phocuswright2022': 'فوكسرايت، 2022',
    'cooper2018': 'كوبر، 2018',
    'unwto2023': 'منظمة السياحة العالمية، 2023',
    'middleton2009': 'ميدلتون وفايال، 2009',
    'alshammari2018': 'الشمري، 2018',
    'kracht2010': 'كراخت وونغ، 2010',
    'standing2014': 'ستاندينغ وآخرون، 2014',
    'iata2023': 'إياتا، 2023',
    'zeithaml2018': 'زايتمل، 2018',
    'law2014': 'لو وآخرون، 2014',
    'porter2008': 'بورتر، 2008',
    'xiang2015': 'شيانغ وآخرون، 2015',
    'amaro2015': 'أمارو وديوغو، 2015',
    'google2022': 'غوغل، 2022',
    'statista2023': 'ستاتيستا، 2023',
    'booking2023': 'بوكينغ، 2023',
    'laudon2020': 'لودون، 2020',
    'albalushi2019': 'البلوشي، 2019',
    'alhammad2021': 'الحماد، 2021',
    'alqahtani2020': 'القحطاني، 2020',
    'euromonitor2023': 'يورومونيتور، 2023',
    'mckinsey2022': 'ماكنزي، 2022',
    'wttc2023': 'مجلس السفر والسياحة العالمي، 2023',
    'bennett2012': 'بينيت، 2012',
    'christensen2016': 'كريستنسن، 2016',
    'booking2022': 'بوكينغ، 2022',
    'kotler2017': 'كوتلر وكيلر، 2017',
    'drucker2015': 'دراكر، 2015',
    'inversini2014': 'إنفرسيني وماسيرو، 2014',
    'wahab2012': 'وهاب، 2012',
    'alnajjar2017': 'النجار، 2017',
}


def resolve_inputs(content, base_dir):
    """Recursively resolve \\input{} commands."""
    def replace_input(match):
        filename = match.group(1)
        if not filename.endswith('.tex'):
            filename += '.tex'
        filepath = os.path.join(base_dir, filename)
        if os.path.exists(filepath):
            with open(filepath, 'r', encoding='utf-8') as f:
                return f.read()
        return match.group(0)

    prev = None
    while prev != content:
        prev = content
        content = re.sub(r'\\input\{([^}]+)\}', replace_input, content)
    return content


def replace_cite(match):
    key = match.group(1)
    name = CITATIONS.get(key, key)
    return f'({name})'


def clean_latex(body):
    """Remove XeLaTeX-specific and visual-only commands, keep structure."""

    # Replace \textenglish{...} → keep content
    body = re.sub(r'\\textenglish\{([^}]*)\}', r'\1', body)

    # Replace \parencite{...} → (Author, Year)
    body = re.sub(r'\\parencite\{([^}]*)\}', replace_cite, body)

    # Remove visual spacing / page commands
    body = re.sub(r'\\vspace\*?\{[^}]*\}', '', body)
    body = re.sub(r'\\vfill', '', body)
    body = re.sub(r'\\newpage', '', body)
    body = re.sub(r'\\thispagestyle\{[^}]*\}', '', body)
    body = re.sub(r'\\pagenumbering\{[^}]*\}', '', body)
    body = re.sub(r'\\markboth\{[^}]*\}\{[^}]*\}', '', body)
    body = re.sub(r'\\addcontentsline\{[^}]*\}\{[^}]*\}\{[^}]*\}', '', body)
    body = re.sub(r'\\blankpage', '', body)
    body = re.sub(r'\\tableofcontents', '', body)
    body = re.sub(r'\\listoftables', '', body)
    body = re.sub(r'\\listoffigures', '', body)
    body = re.sub(r'\\setlength\{[^}]*\}\{[^}]*\}', '', body)
    body = re.sub(r'\\renewcommand\{[^}]*\}\{[^}]*\}', '', body)
    body = re.sub(r'\\centering', '', body)

    # Remove font size commands
    for cmd in ['Huge', 'LARGE', 'Large', 'large', 'normalsize', 'small',
                'footnotesize', 'scriptsize', 'tiny']:
        body = re.sub(rf'\\{cmd}\b', '', body)

    # Remove layout commands
    body = re.sub(r'\\hfill', '', body)
    # Replace \\[0.3cm] style breaks with \\ (keep plain \\ for pandoc)
    body = re.sub(r'\\\\\[[\d.]+cm\]', r'\\\\', body)

    # Remove minipage, flushleft, flushright, titlepage (keep content)
    for env in ['minipage', 'flushright', 'flushleft', 'titlepage']:
        body = re.sub(rf'\\begin\{{{env}\}}(\[[^\]]*\])?\{{[^}}]*\}}', '', body)
        body = re.sub(rf'\\begin\{{{env}\}}(\[[^\]]*\])?', '', body)
        body = re.sub(rf'\\end\{{{env}\}}', '', body)

    # Remove \fcolorbox{rule}{bg}{content} → content
    # Use a function to handle nested braces properly
    def remove_fcolorbox(text):
        pattern = r'\\fcolorbox\{[^}]*\}\{[^}]*\}\{'
        while True:
            m = re.search(pattern, text)
            if not m:
                break
            start = m.start()
            # Find the matching closing brace
            brace_start = m.end() - 1  # position of the opening {
            depth = 1
            pos = m.end()
            while pos < len(text) and depth > 0:
                if text[pos] == '{':
                    depth += 1
                elif text[pos] == '}':
                    depth -= 1
                pos += 1
            # Extract content between braces
            content = text[brace_start + 1:pos - 1]
            text = text[:start] + content + text[pos:]
        return text
    body = remove_fcolorbox(body)

    # Remove \label and \ref
    body = re.sub(r'\\label\{[^}]*\}', '', body)
    body = re.sub(r'الجدول \\ref\{[^}]*\}', 'الجدول التالي', body)
    body = re.sub(r'\\ref\{[^}]*\}', '', body)

    # Handle multirow: \multirow{n}{width}{text} → text
    def remove_multirow(text):
        pattern = r'\\multirow\{[^}]*\}\{[^}]*\}\{'
        while True:
            m = re.search(pattern, text)
            if not m:
                break
            start = m.start()
            depth = 1
            pos = m.end()
            while pos < len(text) and depth > 0:
                if text[pos] == '{':
                    depth += 1
                elif text[pos] == '}':
                    depth -= 1
                pos += 1
            content = text[m.end():pos - 1]
            text = text[:start] + content + text[pos:]
        return text
    body = remove_multirow(body)

    # \cline{...} → \hline (simplified)
    body = re.sub(r'\\cline\{[^}]*\}', r'\\hline', body)

    # Replace \ldots
    body = re.sub(r'\\ldots', '...', body)

    # Remove [H] float placement
    body = re.sub(r'\[H\]', '', body)

    # Remove % comment lines
    body = re.sub(r'^%.*$', '', body, flags=re.MULTILINE)
    # Remove inline comments (careful not to remove \%)
    body = re.sub(r'(?<!\\)%.*$', '', body, flags=re.MULTILINE)

    # printbibliography
    body = re.sub(r'\\printbibliography(\[[^\]]*\])?', '', body)

    # Remove \parbox
    body = re.sub(r'\\parbox\{[^}]*\}\{', '{', body)

    # $\rightarrow$ → →
    body = re.sub(r'\$\\rightarrow\$', '→', body)
    # $\blacktriangleright$ → ►
    body = re.sub(r'\$\\blacktriangleright\$', '►', body)

    # Clean up multiple blank lines
    body = re.sub(r'\n{4,}', '\n\n\n', body)

    return body


def create_reference_docx(output_path):
    """Create a reference DOCX with RTL and Arabic font settings."""
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document()

    # Set default font and RTL for the whole document
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Traditional Arabic'
    font.size = Pt(14)

    # Set RTL for Arabic font
    rpr = style.element.get_or_add_rPr()

    # Set Arabic font (cs = complex script)
    cs_font = OxmlElement('w:rFonts')
    cs_font.set(qn('w:cs'), 'Traditional Arabic')
    cs_font.set(qn('w:ascii'), 'Times New Roman')
    cs_font.set(qn('w:hAnsi'), 'Times New Roman')
    rpr.append(cs_font)

    # Set cs font size
    sz_cs = OxmlElement('w:szCs')
    sz_cs.set(qn('w:val'), '28')  # 14pt = 28 half-points
    rpr.append(sz_cs)

    # Set RTL
    rtl = OxmlElement('w:rtl')
    rpr.append(rtl)

    # Set bidi on paragraph format
    ppr = style.element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    ppr.append(bidi)

    # Set RTL paragraph alignment
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'right')
    ppr.append(jc)

    # Configure heading styles
    for level in range(1, 5):
        style_name = f'Heading {level}'
        if style_name in doc.styles:
            h_style = doc.styles[style_name]
            h_font = h_style.font
            h_font.name = 'Traditional Arabic'
            if level == 1:
                h_font.size = Pt(22)
            elif level == 2:
                h_font.size = Pt(18)
            elif level == 3:
                h_font.size = Pt(16)
            else:
                h_font.size = Pt(14)

            h_rpr = h_style.element.get_or_add_rPr()
            h_cs_font = OxmlElement('w:rFonts')
            h_cs_font.set(qn('w:cs'), 'Traditional Arabic')
            h_cs_font.set(qn('w:ascii'), 'Times New Roman')
            h_cs_font.set(qn('w:hAnsi'), 'Times New Roman')
            h_rpr.append(h_cs_font)

            h_rtl = OxmlElement('w:rtl')
            h_rpr.append(h_rtl)

            # Set cs size
            h_sz_cs = OxmlElement('w:szCs')
            h_sz_cs.set(qn('w:val'), str(h_font.size.pt * 2) if h_font.size else '28')
            h_rpr.append(h_sz_cs)

            h_ppr = h_style.element.get_or_add_pPr()
            h_bidi = OxmlElement('w:bidi')
            h_ppr.append(h_bidi)

            h_jc = OxmlElement('w:jc')
            h_jc.set(qn('w:val'), 'right')
            h_ppr.append(h_jc)

    # Set document-level RTL
    sect_pr = doc.sections[0]._sectPr
    bidi_doc = OxmlElement('w:bidi')
    sect_pr.append(bidi_doc)

    # Set page margins (right=3cm for binding, left=2cm)
    for section in doc.sections:
        section.left_margin = Cm(2)
        section.right_margin = Cm(3)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)

    # Add minimal content so pandoc picks up styles
    p = doc.add_paragraph('.')
    p.clear()

    doc.save(output_path)
    print(f"  Created reference DOCX: {output_path}")


def main():
    print("=" * 60)
    print("  Arabic LaTeX → DOCX Converter")
    print("=" * 60)

    # Step 1: Read and consolidate LaTeX
    print("\n[1/4] Reading and consolidating LaTeX files...")
    main_tex_path = os.path.join(BASE, 'main.tex')
    with open(main_tex_path, 'r', encoding='utf-8') as f:
        main_content = f.read()

    full_content = resolve_inputs(main_content, BASE)

    # Extract body
    doc_match = re.search(
        r'\\begin\{document\}(.*?)\\end\{document\}',
        full_content, re.DOTALL
    )
    if not doc_match:
        print("ERROR: Could not find \\begin{document}...\\end{document}")
        return
    body = doc_match.group(1)

    # Step 2: Clean LaTeX
    print("[2/4] Cleaning LaTeX for pandoc compatibility...")
    body = clean_latex(body)

    # Write cleaned pandoc-compatible LaTeX
    clean_preamble = r"""\documentclass[a4paper,14pt]{extreport}
\usepackage[utf8]{inputenc}
\usepackage{longtable}
\usepackage{booktabs}
\usepackage{multirow}
\usepackage{array}
\usepackage{amssymb}
\usepackage{graphicx}
\usepackage{float}
\usepackage[shortlabels]{enumitem}
\begin{document}
"""

    pandoc_tex = os.path.join(BASE, 'main_pandoc.tex')
    with open(pandoc_tex, 'w', encoding='utf-8') as f:
        f.write(clean_preamble + body + '\n\\end{document}\n')
    print(f"  Written: main_pandoc.tex ({len(body)} chars)")

    # Step 3: Create reference DOCX with RTL/Arabic settings
    print("[3/4] Creating reference DOCX template with Arabic/RTL settings...")
    ref_docx = os.path.join(BASE, 'reference.docx')
    create_reference_docx(ref_docx)

    # Step 4: Run pandoc
    print("[4/4] Converting to DOCX with pandoc...")
    output_docx = os.path.join(BASE, 'main_word.docx')
    cmd = [
        'pandoc',
        pandoc_tex,
        '-f', 'latex',
        '-t', 'docx',
        '-o', output_docx,
        '--reference-doc', ref_docx,
        '-M', 'dir:rtl',
        '-M', 'lang:ar',
        '--wrap=none',
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  pandoc stderr: {result.stderr}")
        # Try again without some options if it fails
        print("  Retrying with simpler options...")
        cmd2 = [
            'pandoc',
            pandoc_tex,
            '-f', 'latex',
            '-t', 'docx',
            '-o', output_docx,
            '--reference-doc', ref_docx,
        ]
        result2 = subprocess.run(cmd2, capture_output=True, text=True)
        if result2.returncode != 0:
            print(f"  pandoc error: {result2.stderr}")
            return
    
    if result.stderr:
        # Show warnings but continue
        warnings = result.stderr.strip().split('\n')
        if len(warnings) > 5:
            print(f"  pandoc: {len(warnings)} warnings (showing first 5):")
            for w in warnings[:5]:
                print(f"    {w}")
        else:
            for w in warnings:
                print(f"  pandoc warning: {w}")

    if os.path.exists(output_docx):
        size_mb = os.path.getsize(output_docx) / (1024 * 1024)
        print(f"\n{'=' * 60}")
        print(f"  SUCCESS! Created: main_word.docx ({size_mb:.2f} MB)")
        print(f"  Location: {output_docx}")
        print(f"{'=' * 60}")
    else:
        print("\n  ERROR: Output file was not created.")


if __name__ == '__main__':
    main()
