#!/usr/bin/env python3
"""
BOB Autowas — Blog Converter
=============================
Leest alle .docx bestanden uit WebsiteBA/Blogs/ en verwerkt ze automatisch
in BOB_Blog.html.

Gebruik:
  python blog_converter.py

Vereist: pip install python-docx
"""

import os
import re
import json
import sys
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── PADEN ──────────────────────────────────────────────────────────
SCRIPT_DIR  = Path(__file__).parent
BLOGS_DIR   = SCRIPT_DIR / 'Blogs'
BLOG_HTML   = SCRIPT_DIR / 'BOB_Blog.html'

# ── HELPERS ────────────────────────────────────────────────────────

def parse_meta_line(text):
    """Haal key: value op uit een metaregel."""
    if ':' not in text:
        return None, None
    key, _, val = text.partition(':')
    return key.strip().upper(), val.strip()


def para_to_html(para):
    """Zet een Word-alinea om naar HTML."""
    style = para.style.name if para.style else ''
    text = para.text.strip()
    if not text:
        return ''

    # Heading niveaus
    if 'Heading 1' in style:
        return f'<h2>{escape(text)}</h2>'
    if 'Heading 2' in style:
        return f'<h2>{escape(text)}</h2>'
    if 'Heading 3' in style:
        return f'<h3>{escape(text)}</h3>'

    # Bullets / opsomming
    if para.style and 'List' in para.style.name:
        return f'<li>{run_to_html(para)}</li>'

    # Citaat (ingesprongen / quote)
    fmt = para.paragraph_format
    if fmt.left_indent and fmt.left_indent > 200000:  # EMU
        return f'<blockquote>{run_to_html(para)}</blockquote>'

    return f'<p>{run_to_html(para)}</p>'


def run_to_html(para):
    """Verwerk runs (vet, cursief) naar HTML."""
    parts = []
    for run in para.runs:
        t = escape(run.text)
        if not t:
            continue
        if run.bold and run.italic:
            t = f'<strong><em>{t}</em></strong>'
        elif run.bold:
            t = f'<strong>{t}</strong>'
        elif run.italic:
            t = f'<em>{t}</em>'
        parts.append(t)
    return ''.join(parts) or escape(para.text)


def escape(s):
    return (s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def wrap_lists(html_parts):
    """Wrap aaneengesloten <li> items in <ul>."""
    result = []
    in_list = False
    for part in html_parts:
        if part.startswith('<li>'):
            if not in_list:
                result.append('<ul>')
                in_list = True
        else:
            if in_list:
                result.append('</ul>')
                in_list = False
        result.append(part)
    if in_list:
        result.append('</ul>')
    return result


def docx_to_blog(docx_path):
    """Lees een .docx en geef een blog-dict terug."""
    doc = Document(docx_path)
    slug = docx_path.stem  # bestandsnaam zonder .docx

    meta = {
        'SLUG': slug,
        'CATEGORIE': 'tips',
        'DATUM': '',
        'LEESTIJD': '3 min',
        'FOTO 1 (BESTANDSNAAM)': f'{slug}-1.jpg',
        'FOTO 2 (BESTANDSNAAM)': f'{slug}-2.jpg',
    }
    title = ''
    excerpt = ''
    body_parts = []

    # Fasen: META, SAMENVATTING, BODY
    phase = 'META'
    prev_was_excerpt_label = False

    for para in doc.paragraphs:
        text = para.text.strip()
        style = para.style.name if para.style else ''

        # Detecteer sectie-labels
        if text in ('BLOGTEKST (schrijf hieronder uw blog):', 'BLOGTEKST'):
            phase = 'BODY'
            continue
        if 'KORTE SAMENVATTING' in text:
            phase = 'SAMENVATTING'
            continue
        if 'BLOGTITEL' in text:
            continue
        if text.startswith('INSTRUCTIES:') or 'BOB AUTOWAS' in text:
            continue
        if text.startswith('BLOG INFORMATIE'):
            continue
        if text.startswith('OPGEMAAKTE ELEMENTEN'):
            break  # voettekst — stop

        if phase == 'META':
            key, val = parse_meta_line(text)
            if key and key in meta and val:
                meta[key] = val
            # Titel = Heading 1 in meta-fase of daarna
            if 'Heading 1' in style and text:
                title = text
                phase = 'SAMENVATTING_LABEL'
            continue

        if phase == 'SAMENVATTING_LABEL':
            if 'KORTE SAMENVATTING' in text:
                phase = 'SAMENVATTING'
            elif 'Heading 1' in style and text:
                title = text
            continue

        if phase == 'SAMENVATTING':
            if text and not text.startswith('BLOGTEKST'):
                excerpt = text
                phase = 'BODY_LABEL'
            continue

        if phase == 'BODY_LABEL':
            if 'BLOGTEKST' in text:
                phase = 'BODY'
            continue

        if phase == 'BODY':
            html_part = para_to_html(para)
            if html_part:
                body_parts.append(html_part)

    body_parts = wrap_lists(body_parts)
    inhoud = '\n'.join(body_parts)

    # Foto paden (relatief t.o.v. de HTML pagina)
    foto1_name = meta.get('FOTO 1 (BESTANDSNAAM)', f'{slug}-1.jpg')
    foto2_name = meta.get('FOTO 2 (BESTANDSNAAM)', f'{slug}-2.jpg')
    foto1 = f'Blogs/{foto1_name}'
    foto2 = f'Blogs/{foto2_name}'

    return {
        'id':        meta.get('SLUG', slug),
        'titel':     title or slug,
        'excerpt':   excerpt,
        'categorie': meta.get('CATEGORIE', 'tips').lower(),
        'datum':     meta.get('DATUM', ''),
        'leestijd':  meta.get('LEESTIJD', '3 min'),
        'foto1':     foto1,
        'foto2':     foto2,
        'inhoud':    inhoud,
    }


def js_string(s):
    """Escape voor gebruik in een JS single-quoted string."""
    return s.replace('\\', '\\\\').replace("'", "\\'").replace('\n', '\\n').replace('\r', '')


def blogs_to_js(blogs):
    """Genereer de var BLOGS = [...] JS array."""
    lines = ['var BLOGS = [']
    for i, b in enumerate(blogs):
        comma = '' if i == len(blogs) - 1 else ','
        lines.append('  {')
        lines.append(f"    id: '{js_string(b['id'])}',")
        lines.append(f"    titel: '{js_string(b['titel'])}',")
        lines.append(f"    excerpt: '{js_string(b['excerpt'])}',")
        lines.append(f"    categorie: '{js_string(b['categorie'])}',")
        lines.append(f"    datum: '{js_string(b['datum'])}',")
        lines.append(f"    leestijd: '{js_string(b['leestijd'])}',")
        lines.append(f"    foto1: '{js_string(b['foto1'])}',")
        lines.append(f"    foto2: '{js_string(b['foto2'])}',")
        inhoud_escaped = js_string(b['inhoud'])
        lines.append(f"    inhoud: '{inhoud_escaped}'")
        lines.append(f'  }}{comma}')
    lines.append('];')
    return '\n'.join(lines)


def update_blog_html(blogs):
    """Vervang var BLOGS = [...] in BOB_Blog.html."""
    with open(BLOG_HTML, encoding='utf-8') as f:
        html = f.read()

    new_js = blogs_to_js(blogs)

    # Vervang bestaande BLOGS array
    pattern = re.compile(r'var BLOGS = \[.*?\];', re.DOTALL)
    if pattern.search(html):
        html = pattern.sub(new_js, html, count=1)
    else:
        print("WAARSCHUWING: var BLOGS niet gevonden in BOB_Blog.html")
        return False

    with open(BLOG_HTML, 'w', encoding='utf-8') as f:
        f.write(html)
    return True


# ── MAIN ───────────────────────────────────────────────────────────

def main():
    if not BLOGS_DIR.exists():
        print(f"Map '{BLOGS_DIR}' bestaat niet. Aanmaken...")
        BLOGS_DIR.mkdir(parents=True)

    if not BLOG_HTML.exists():
        print(f"FOUT: '{BLOG_HTML}' niet gevonden.")
        sys.exit(1)

    # Zoek alle .docx bestanden in Blogs/
    docx_files = sorted(BLOGS_DIR.glob('*.docx'))

    if not docx_files:
        print(f"Geen .docx bestanden gevonden in {BLOGS_DIR}")
        print("Tip: Zet uw blog .docx bestanden in de map Blogs/")
        sys.exit(0)

    blogs = []
    for docx_path in docx_files:
        print(f"Verwerken: {docx_path.name} ...", end=' ')
        try:
            blog = docx_to_blog(docx_path)
            blogs.append(blog)
            print(f"OK  ({blog['titel'][:50]})")
        except Exception as e:
            print(f"FOUT: {e}")

    if not blogs:
        print("Geen blogs verwerkt.")
        sys.exit(1)

    print(f"\n{len(blogs)} blog(s) verwerkt. BOB_Blog.html bijwerken...")
    ok = update_blog_html(blogs)

    if ok:
        print("✓ BOB_Blog.html succesvol bijgewerkt!")
        print()
        print("Blogs op de pagina:")
        for b in blogs:
            print(f"  - [{b['categorie']}] {b['titel']} ({b['datum']})")
    else:
        print("✗ Bijwerken mislukt.")
        sys.exit(1)


if __name__ == '__main__':
    main()
