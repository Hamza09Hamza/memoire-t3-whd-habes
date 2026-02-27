#!/usr/bin/env python3
"""Inspect the front_pages.docx structure."""
from docx import Document

doc = Document('front_pages.docx')
body = doc.element.body
ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
dns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
wns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

children = list(body)
for i, child in enumerate(children):
    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
    text = ''
    if tag == 'p':
        for t in child.iter(f'{ns}t'):
            text += (t.text or '')
        text = text[:80]
    elif tag == 'tbl':
        text = '[TABLE]'
    elif tag == 'sectPr':
        text = '[SECTION]'
    
    # Check for drawings
    from lxml import etree
    drawings = list(child.iter(f'{wns}drawing'))
    blips = list(child.iter(f'{dns}blip'))
    img = f' [DRAW:{len(drawings)} BLIP:{len(blips)}]' if drawings or blips else ''
    
    print(f'[{i:3d}] {tag:10s} | {text}{img}')

print(f'\nTotal body children: {len(children)}')

# Check images in the package
import zipfile
import os
docx_path = 'front_pages.docx'
with zipfile.ZipFile(docx_path, 'r') as z:
    images = [f for f in z.namelist() if f.startswith('word/media/')]
    print(f'\nImages in DOCX: {len(images)}')
    for img in images:
        info = z.getinfo(img)
        print(f'  {img} ({info.file_size} bytes)')
