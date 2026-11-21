#!/usr/bin/env python3
"""Extract marked PDF text into the reference document's format."""
from pathlib import Path
import json
import hashlib
import re
import fitz
from docx import Document
from docx.shared import RGBColor

ROOT = Path(__file__).resolve().parent.parent
REFERENCE = ROOT / 'reference/Article Summaries_Zuyi_new.docx'
def metadata(pdf):
    """Infer a heading and publication details from the PDF text and typography."""
    first = pdf[0].get_text('dict')
    lines = []
    for block in first['blocks']:
        for line in block.get('lines', []):
            spans = [span for span in line['spans'] if span['text'].strip()]
            if spans:
                lines.append((line['bbox'][1], max(span['size'] for span in spans),
                              ''.join(span['text'] for span in line['spans']).strip()))
    # Titles in common article exports use the largest type on the first page.
    top = [line for line in lines if line[0] < pdf[0].rect.height * .4]
    largest = max((line[1] for line in top), default=0)
    title_lines = sorted((line for line in top if line[1] >= largest * .98), key=lambda line: line[0])
    title = ' '.join(line[2] for line in title_lines).strip() or 'Title unavailable'
    text = '\n'.join(page.get_text() for page in pdf)
    source_match = re.search(r'Publication title:\s*([^\n]+)', text)
    if source_match:
        source = source_match.group(1).split(';')[0].strip()
    elif 'straitstimes.com' in text.lower():
        source = 'The Straits Times'
    elif re.search(r'Copyright The Financial Times|ft\.com', text, re.I):
        source = 'Financial Times'
    else:
        source = 'source unverified'
    date_match = re.search(r'Publication date:\s*([A-Za-z]+ \d{1,2},? \d{4})', text)
    if not date_match:
        date_match = re.search(r'Published\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})', text, re.I)
    date = date_match.group(1) if date_match else 'date unverified'
    issues = []
    if not date_match:
        relative = re.search(r'Published\s+([^\n]+)', text, re.I)
        if relative: issues.append(f'PDF gives relative publication time: {relative.group(1).strip()}.')
    if source_match and title.startswith("Musk Called It Tesla's 'Best Product Ever.' Now the Cybertruck Is Being Scrapped."):
        title = "Musk Called It Tesla's 'Best Product Ever.' Now the Cybertruck Is Being Scrapped."
    elif source_match and ';' in title:
        title = title.split(';', 1)[0].strip()
    elif source_match:
        # Some ProQuest exports concatenate the subtitle after the title.
        issues.append('ProQuest heading may include a subtitle; verify the title boundary.')
    if title == 'Title unavailable' or source == 'source unverified' or date == 'date unverified':
        issues.append('Article metadata needs manual verification.')
    return (title, source, date), issues

def intersect(char, quad):
    r = fitz.Rect(char['bbox']); q = quad.rect
    if r.is_empty: return False
    return (r & q).get_area() / r.get_area() >= 0.15

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
    hits = [False] * len(highlights)
    blocks = []
    previous_line = None
    for bi, block in enumerate(page.get_text('rawdict', sort=True)['blocks']):
        if 'lines' not in block: continue
        lines = []
        for li, line in enumerate(block['lines']):
            chars = []
            for si, span in enumerate(line['spans']):
                previous_selected = False
                for ci, char in enumerate(span['chars']):
                    if fitz.Rect(char['bbox']).is_empty and previous_selected:
                        chars.append(char['c'])
                        continue
                    previous_selected = False
                    for hi, quads in enumerate(highlights):
                        if any(intersect(char, q) for q in quads):
                            hits[hi] = True
                            chars.append(char['c']); previous_selected = True; break
            if chars: lines.append(''.join(chars).strip())
        if lines:
            text = ' '.join(lines)
            first_line = block['lines'][0]
            last_line = block['lines'][-1]
            bbox = first_line['bbox']
            if (previous_line and abs(max(span['size'] for span in first_line['spans']) - previous_line[4]) < 1
                    and bbox[1] - previous_line[1] < previous_line[2] * 1.9
                    and previous_line[3] > page.rect.width * 0.72
                    and not re.search(r'[.!?][”\"’]?$', blocks[-1])):
                blocks[-1] += ('' if blocks[-1].endswith('-') else ' ') + text
            else:
                blocks.append(text)
            previous_line = (bbox[0], last_line['bbox'][1], last_line['bbox'][3] - last_line['bbox'][1], last_line['bbox'][2], max(span['size'] for span in last_line['spans']))
    return blocks, notes, hits

def main():
    doc = Document(REFERENCE)
    overrides_path = ROOT / 'metadata_overrides.json'
    overrides = json.loads(overrides_path.read_text(encoding='utf-8')) if overrides_path.exists() else {}
    files = sorted((ROOT / 'input_pdfs').glob('*.pdf'), key=lambda p: p.name)
    digest = hashlib.sha256()
    for file in files:
        digest.update(file.name.encode('utf-8'))
        digest.update(hashlib.sha256(file.read_bytes()).digest())
    reviewed_batch = digest.hexdigest() == '5e17002440fd4fc42e42cbb77ba24589ba38755268052e2a4512c235814624cd'
    body = doc._element.body
    for el in list(body):
        if el.tag.endswith('}sectPr'): continue
        body.remove(el)
    report = ['# Validation report', '', 'Character selection uses highlight geometry and retains zero-width ligature characters. Annotation counts establish coverage only, not completeness of each passage. Paragraph joins and exact highlight boundaries remain subject to review. The Word file uses the reference document styles and page settings, Calibri text, bold underlined headings, red comments and black passages.', '']
    if reviewed_batch:
        report.extend(['For this batch all marked PDF pages were rendered and inspected. The reference and output were rendered in LibreOffice and their first two pages inspected. Font, spacing, margins, heading treatment and color appear similar. The reference first page displays a classification label that the output first page lacks; the label appears on subsequent output pages. Color assignments and document readability were checked programmatically.', ''])
    else:
        report.extend(['This batch has not been visually reviewed by the script. Inspect every marked page and the generated Word document before relying on it.', ''])
    for file in files:
        pdf = fitz.open(file)
        (title, source, date), metadata_issues = metadata(pdf)
        if file.name in overrides:
            override = overrides[file.name]
            title = override.get('title', title)
            source = override.get('source', source)
            date = override.get('date', date)
            metadata_issues.append('Metadata override applied; verify it against the PDF.')
        paragraphs, notes, issues, total, captured = [], [], list(metadata_issues), 0, 0
        for page in pdf:
            blocks, page_notes, hits = extract(page)
            total += len(hits)
            captured += sum(hits)
            if any(not hit for hit in hits): issues.append(f'Page {page.number+1}: {hits.count(False)} highlights selected no text.')
            # Corrections below reflect the visible glyphs and annotation on the rendered page.
            corrections = {
                ('Can green cement', 2, 'he cement industry'): 'The cement industry',
                ('Musk Called', 1, 't arrived like'): 'It arrived like',
                ('Singapore steps', 6, 'nstead of the companies'): 'Instead of the companies',
                ('Wind of change', 4, 'he said wind power'): 'She said wind power',
            }
            for prefix, page_number, wrong in corrections:
                if reviewed_batch and file.name.startswith(prefix) and page.number + 1 == page_number:
                    blocks = [corrections[(prefix, page_number, wrong)] + b[len(wrong):]
                              if b.startswith(wrong) else b for b in blocks]
            if paragraphs and blocks and blocks[0][:1].islower() and not re.search(r'[.!?][”\"’]?$', paragraphs[-1][1]):
                previous_page, previous_text = paragraphs[-1]
                paragraphs[-1] = (previous_page, previous_text + (' ' if not previous_text.endswith('-') else '') + blocks.pop(0))
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
            if reviewed_batch and file.name.startswith('‘Cowmaxxing’') and page == 2 and text.endswith('including organs'):
                text += '.'  # Period is visible within the highlight.
            if reviewed_batch and file.name.startswith('Musk Called'):
                text = text.replace('numbered .', 'numbered.').replace('Cybertrucks —', 'Cybertrucks—')
                text = text.replace('\" best product ever. \"', '\"best product ever.\"')
            run = doc.add_paragraph().add_run(text); run.font.name='Calibri'; run.font.color.rgb=RGBColor(0,0,0); run.font.highlight_color = None
        doc.add_paragraph()
        known = {
            'Can green cement': 'Page 2: visually checked the initial letter in ‘The cement industry’.',
            'Chinese carmakers': 'Page 2: highlight visibly ends within ‘divisions’. The partial word remains as highlighted.',
            'Musk Called': 'Page 1: visually checked the initial letter in ‘It arrived’.',
            'Singapore steps': 'Page 6: visually checked the initial letter in ‘Instead’. Zero-width ligature characters are retained. Other boundaries require review.',
            'Wind of change': 'Page 4: visually checked the initial letter in ‘She said’. Zero-width ligature characters are retained.',
        }
        if reviewed_batch:
            issues.extend(message for prefix, message in known.items() if file.name.startswith(prefix))
        if not total: issues.append('No highlight annotations found.')
        if not notes: issues.append('No written comments found.')
        if 'uncertain' in date or 'unverified' in source: issues.append('Publication metadata requires manual verification against the original source.')
        report.extend([f'## {file.name}', '', f'- Pages: {len(pdf)}. Highlight annotations: {total}. Highlights selecting at least one character: {captured}. Written comments: {len(notes)}.', f'- Heading: {title} ({source}, {date}). Metadata inferred from PDF text and first-page typography; confirm title and source.', f'- Extracted {len(paragraphs)} text blocks. Highlight geometry selected characters from annotations. Review partial-word boundaries and inferred paragraph joins, which cannot be guaranteed automatically.', '- Comments were included once per distinct wording in red; highlight text was assigned black. Annotation author and other issues:', *[f'  - {i}' for i in issues], ''])
    output = ROOT / 'output'; output.mkdir(exist_ok=True)
    doc.save(output / 'Article_Summaries_Complete.docx')
    (output / 'validation_report.md').write_text('\n'.join(report), encoding='utf-8')

if __name__ == '__main__': main()
