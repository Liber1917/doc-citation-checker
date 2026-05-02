# Doc Citation Checker

A WorkBuddy skill that verifies citations in academic documents **without requiring a `.bib` file**.

Supports `.docx`, `.md`, `.txt`, and `.tex` documents. Detects AI-hallucinated references using CrossRef, Semantic Scholar, and OpenAlex APIs.

## Features

- ✅ Works directly with Word, Markdown, plain text, and LaTeX files
- ✅ No `.bib` file required
- ✅ Detects hallucinated citations (AI-fabricated references)
- ✅ Verifies against CrossRef, Semantic Scholar, and OpenAlex
- ✅ Title similarity matching to avoid false positives
- ✅ LaTeX: maps `\cite{key}` to `\bibitem` content for accurate verification

## Installation

Clone this repository to your WorkBuddy skills directory:

```bash
git clone https://github.com/YOUR_USERNAME/doc-citation-checker.git ~/.workbuddy/skills/doc-citation-checker
```

Install dependencies:

```bash
pip install requests python-docx
```

## Usage

### Check a single file

```bash
python ~/.workbuddy/skills/doc-citation-checker/scripts/check_citations.py --file paper.docx
```

### Check a directory

```bash
python ~/.workbuddy/skills/doc-citation-checker/scripts/check_citations.py --dir ./papers/
```

### Save report to JSON

```bash
python ~/.workbuddy/skills/doc-citation-checker/scripts/check_citations.py --file thesis.tex --output report.json
```

### Adjust similarity threshold

```bash
python ~/.workbuddy/skills/doc-citation-checker/scripts/check_citations.py --file paper.md --threshold 0.4
```

## Supported File Types

| Format | Extraction Method |
|--------|------------------|
| `.docx` | python-docx paragraph extraction |
| `.md` / `.txt` | Regex: DOI, bibliography section, author-year |
| `.tex` | `\bibitem` map + `\cite{key}` resolution |

## Output Example

```
============================================================
CITATION VERIFICATION REPORT
============================================================
File: thesis.tex
Total citations: 15
  ✅ Valid:         13
  ⚠️  Suspicious:    0
  ❌ Hallucinated:  1
  ❓ Unverifiable:  1

❌ HALLUCINATED CITATIONS (not found in any database):
  - [fake2099] FakeAuthor. ``Quantum time travel.'' Journal of Imaginary Science.
============================================================
```

## How It Works

1. **Extract** citations from the document using format-specific parsers
2. **Resolve** `.tex` `\cite{key}` references to their `\bibitem` text
3. **Verify** each citation via DOI (CrossRef) or title search (CrossRef + Semantic Scholar)
4. **Match** search results using Jaccard title similarity to avoid false positives
5. **Report** valid, hallucinated, and unverifiable citations

## Testing

```bash
python scripts/check_citations.py --file tests/sample_paper.tex
python scripts/check_citations.py --file tests/sample_paper.md
```

## License

MIT
