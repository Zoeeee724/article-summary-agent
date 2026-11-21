#!/usr/bin/env python3
"""Extract marked PDF text into the reference document's format."""
from pathlib import Path
from collections import defaultdict
import re
import fitz
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor

ROOT = Path(__file__).resolve().parent.parent
REFERENCE = ROOT / 'reference/Article Summaries_Zuyi_new.docx'
# Metadata transcribed from the PDFs. Extend for subsequent batches.
METADATA = {
 'Can green cement': ('Can green cement decarbonise the data centre building boom?', 'Financial Times', '12 Nov 2026'),
 'Chinese carmakers': ('Chinese carmakers offer lifeline to Europe’s struggling steel industry', 'Financial Times', '12 Nov 2026'),
 'Microsoft outage': ('Microsoft outage exposes global reliance on core tech providers', 'Financial Times', 'date uncertain; PDF says Published 11 HOURS AGO'),
 "Musk Called": ("Musk Called It Tesla's 'Best Product Ever.' Now the Cybertruck Is Being Scrapped.", 'Wall Street Journal (Online)', '6 Nov 2026'),
 'Singapore steps': ('Singapore steps up carbon market push with grants, new initiatives at COP31', 'The Straits Times', '13 Nov 2026'),
 'Supermarket-Fridge': ('Supermarket-Fridge Hacks Show Risk of Linking Everything to the Internet', 'Wall Street Journal (Online)', '9 Nov 2026'),
 'Wind of change': ('Wind of change: Laos project shows promise of connecting S-E Asia to green power', 'The Straits Times', '9 Nov 2026'),
 '‘Cowmaxxing’': ('‘Cowmaxxing’: Meat companies use every bit of the animal to fuel wellness boom', 'Financial Times', '8 Nov 2026'),
}

def intersect(char, quad):
    r = fitz.Rect(char['bbox']); q = quad.rect
    if r.is_empty: return False
    return (r & q).get_area() / r.get_area() >= 0.5

def extract(page):
    highlights = []
    notes = []
    for annotation in page.annots() or []:
        info = annotation.info
        if info.get('content', '').strip():
            notes.append((annotation.type[1], info['content'], info.get('title', '')))
        if annotation.type[1] == 'Highlight':
            v = annotation.vertices or []
            highlights.append([fitz.Quad(v[i:i+4]) for i in range(0, len(v), 4)])
    selected = set()
    hits = [False] * len(highlights)
    blocks = []
    for bi, block in enumerate(page.get_text('rawdict', sort=True)['blocks']):
        if 'lines' not in block: continue
        lines = []
        for li, line in enumerate(block['lines']):
            chars = []
            for si, span in enumerate(line['spans']):
                for ci, char in enumerate(span['chars']):
                    for hi, quads in enumerate(highlights):
                        if any(intersect(char, q) for q in quads):
                            selected.add((bi, li, si, ci)); hits[hi] = True
                            chars.append(char['c']); break
            if chars: lines.append(''.join(chars).strip())
        if lines:
            # Line wrapping does not imply a paragraph break.
            blocks.append(' '.join(lines))
    return blocks, notes, hits

def main():
    doc = Document(REFERENCE)
    body = doc._element.body
    for el in list(body):
        if el.tag.endswith('}sectPr'): continue
        body.remove(el)
    report = ['# Validation report', '', 'Each PDF was opened and annotations were enumerated. Character selection uses highlight geometry; extracted text requires manual comparison with the rendered pages, especially at boundaries and line breaks. Counts show annotation coverage, not an independent proof of textual accuracy.', '']
    for file in sorted((ROOT / 'input_pdfs').glob('*.pdf'), key=lambda p: p.name):
        meta = next((v for k,v in METADATA.items() if file.name.startswith(k)), None)
        if meta is None: meta = (file.stem, 'source unverified', 'date unverified')
        title, source, date = meta
        pdf = fitz.open(file)
        paragraphs, notes, issues, total, captured = [], [], [], 0, 0
        for page in pdf:
            blocks, page_notes, hits = extract(page)
            total += len(hits)
            captured += sum(hits)
            if any(not hit for hit in hits): issues.append(f'Page {page.number+1}: {hits.count(False)} highlights selected no text.')
            paragraphs.extend((page.number+1, b) for b in blocks)
            notes.extend((page.number+1, kind, text, author) for kind, text, author in page_notes)
        heading = doc.add_paragraph()
        run = heading.add_run(f'{title} ({source}, {date})'); run.bold = True; run.underline = True; run.font.name = 'Calibri'; run.font.color.rgb = RGBColor(0,0,0)
        unique_notes = set()
        for page, kind, text, author in notes:
            if text in unique_notes: continue
            unique_notes.add(text)
            run = doc.add_paragraph().add_run(text); run.font.name='Calibri'; run.font.color.rgb=RGBColor(255,0,0)
            if not author: issues.append(f'Page {page}: comment author is blank; authorship unverified.')
        for page, text in paragraphs:
            run = doc.add_paragraph().add_run(text); run.font.name='Calibri'; run.font.color.rgb=RGBColor(0,0,0); run.font.highlight_color = None
        doc.add_paragraph()
        known = {
            'Can green cement': 'Page 2: extracted “he cement industry” appears to lack its initial letter.',
            'Chinese carmakers': 'Page 2: highlight ends with the fragment “However, di”; inspect the highlight boundary.',
            'Musk Called': 'Page 1: extracted “t arrived” appears to lack its initial letter.',
            'Singapore steps': 'Pages 2–6: missing letters appear throughout the PDF text layer, including “ofce” and “nstead”. Inspect rendered pages.',
            'Wind of change': 'Page 4: extracted “he said” may be missing its initial letter; other words have missing letters, including “benefts”.',
        }
        issues.extend(message for prefix, message in known.items() if file.name.startswith(prefix))
        if not total: issues.append('No highlight annotations found.')
        if not notes: issues.append('No written comments found.')
        if 'uncertain' in date or 'unverified' in source: issues.append('Publication metadata requires manual verification against the original source.')
        report.extend([f'## {file.name}', '', f'- Pages: {len(pdf)}. Highlight annotations: {total}. Highlights selecting at least one character: {captured}. Written comments: {len(notes)}.', f'- Heading: {title} ({source}, {date}). Metadata transcribed from PDF; relative publication date flagged where applicable.', f'- Extracted {len(paragraphs)} text blocks. Manual visual review of every highlight, reading order, paragraph breaks, punctuation, duplicates, and clipping remains necessary.', '- Comments were included once per distinct wording in red; highlight text was assigned black. Annotation author and other issues:', *[f'  - {i}' for i in issues], ''])
    output = ROOT / 'output'; output.mkdir(exist_ok=True)
    doc.save(output / 'Article_Summaries_Complete.docx')
    (output / 'validation_report.md').write_text('\n'.join(report), encoding='utf-8')

if __name__ == '__main__': main()
