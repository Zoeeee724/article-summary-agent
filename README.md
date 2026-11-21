# Article summary agent

Extract PDF highlights and comments into a Word document using `reference/Article Summaries_Zuyi_new.docx` as the formatting reference. The script copies marked text and does not generate summaries.

## Run

Install Python dependencies with `python -m pip install -r requirements.txt`. Place the PDFs in `input_pdfs/`, keeping the reference document in `reference/`. Before running, add each new article's title, source and original publication date to `METADATA` in `scripts/process_articles.py`, checking them against the PDF. Unlisted PDFs use their filename and are flagged for metadata review.

Run `python scripts/process_articles.py`. The script replaces `output/Article_Summaries_Complete.docx` and `output/validation_report.md`. Review the report and compare every highlighted passage and comment against its PDF, including paragraph boundaries and uncertain metadata. PDF annotation geometry and extraction may require manual correction.
