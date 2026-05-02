---
name: doc-citation-checker
description: >
  Check citations in documents (.docx, .md, .txt, .tex) for hallucinations and validity without requiring a .bib file.
  Use when the user asks to verify citations, check references, detect AI hallucinations in bibliography, or validate academic sources.
  Supports extracting citations directly from document text and verifying them against CrossRef, Semantic Scholar, and OpenAlex APIs.
agent_created: true
---

# Doc Citation Checker

Check citations in academic documents for hallucinations and validity. This skill works with `.docx`, `.md`, `.txt`, and `.tex` files without requiring a `.bib` file.

## When to Use

Trigger this skill when the user:
- Asks to verify citations or references in a document
- Wants to check for AI hallucinated citations
- Needs to validate academic sources in a paper
- Mentions citation checking, reference validation, or bibliography verification
- Says "检查引用", "验证参考文献", "检测幻觉引用", "citation check"

## Supported File Types

- **`.docx`** - Word documents (uses python-docx)
- **`.md`** - Markdown files
- **`.txt`** - Plain text files
- **`.tex`** - LaTeX files (extracts \cite{} and \bibitem{})

## Workflow

### Step 1: Install Dependencies

Check if required Python packages are installed:

```bash
python -c "import requests, re, json"
```

If missing, install:
```bash
pip install requests python-docx
```

### Step 2: Run the Citation Checker

Execute the citation checker script:

```bash
python "~/.workbuddy/skills/doc-citation-checker/scripts/check_citations.py" --file "<document-path>"
```

Optional arguments:
- `--file <path>` - Single file to check
- `--dir <path>` - Directory to scan for documents
- `--output <path>` - Save report to JSON file
- `--threshold <float>` - Similarity threshold (default: 0.85)

### Step 3: Interpret Results

The script outputs:
- **Valid citations** - Citations found in academic databases
- **Suspicious citations** - Citations with mismatched metadata
- **Hallucinated citations** - Citations not found in any database

## Citation Extraction Patterns

The script automatically extracts citations using these patterns:

### From .docx / .md / .txt:
- Numbered references: `[1]`, `[2]`, etc.
- Author-year: `(Smith, 2020)`, `[Smith, 2020]`
- DOI patterns: `doi:10.xxxx/xxxxx`
- Bibliography sections: Lines starting with numbers orIndentation

### From .tex:
- `\cite{key}` commands
- `\bibitem{key}` entries
- `\begin{thebibliography}` environments

## Verification Databases

Citations are verified against:
1. **CrossRef** - DOI resolution and metadata
2. **Semantic Scholar** - Academic paper database
3. **OpenAlex** - Open academic catalog

## Output Format

The script generates a JSON report:

```json
{
  "file": "paper.docx",
  "total_citations": 15,
  "valid": [...],
  "suspicious": [...],
  "hallucinated": [...],
  "summary": {
    "valid_count": 10,
    "suspicious_count": 3,
    "hallucinated_count": 2
  }
}
```

## Common Issues

### No citations found
- Check if the document has a bibliography section
- Verify citations are in a recognizable format
- Try manually copying citations to a .txt file

### API rate limits
- The script includes delays between API calls
- For large documents, use `--output` to save progress
- Consider running overnight for 50+ citations

### False positives
- Some legitimate citations may not be in databases (preprints, non-English, very recent)
- Manually verify suspicious results before concluding hallucination
