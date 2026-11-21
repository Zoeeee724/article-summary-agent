# Article summary agent

Extract highlighted PDF passages and written annotations into one Word document using `reference/Article Summaries_Zuyi_new.docx` for styles and page settings. The script copies marked content and does not generate summaries.

Upload the next batch to `input_pdfs/` and ask Codex to process it. The script processes every PDF in that folder in filename order, including PDFs without annotations. Move completed PDFs out of the input folder before a new batch if they should not be included again.

Install dependencies with `python -m pip install -r requirements.txt`, then run `python scripts/process_articles.py`. The script replaces `output/Article_Summaries_Complete.docx` and `output/validation_report.md`. It infers titles from first-page typography and sources and dates from PDF text, and flags missing information. The heuristic covers common Financial Times, Straits Times and ProQuest exports; other layouts may require correction. Scanned PDFs without a text layer need OCR and manual verification.

Review the report and compare highlights, comments, metadata and paragraph boundaries with the PDF pages. Annotation geometry can cut through letters or words and extraction cannot reliably determine author intent. Preserve partially highlighted text and flag ambiguous boundaries. If metadata cannot be established automatically, an optional `metadata_overrides.json` at the repository root can map the exact PDF filename to an object containing any of `title`, `source`, and `date`. The report records when an override is used.
