#!/usr/bin/env python3
"""
Doc Citation Checker - Verify citations in documents without .bib file
Supports .docx, .md, .txt, .tex files
"""

import argparse
import json
import re
import sys
import time
from pathlib import Path
from typing import Dict, List, Any, Optional

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


def extract_citations_from_docx(file_path: str) -> List[Dict[str, Any]]:
    """Extract citations from .docx file"""
    if not HAS_DOCX:
        print("Warning: python-docx not installed. Install with: pip install python-docx")
        return []
    
    citations = []
    doc = Document(file_path)
    
    # Extract text from all paragraphs
    full_text = "\n".join([p.text for p in doc.paragraphs])
    
    # Also check tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                full_text += "\n" + cell.text
    
    # Extract citations using various patterns
    citations.extend(extract_citations_from_text(full_text))
    
    return citations


def extract_citations_from_tex(file_path: str) -> List[Dict[str, Any]]:
    """Extract citations from .tex file, mapping \\cite{key} to \\bibitem content"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Step 1: Build a key -> raw text map from \bibitem entries
    bibitem_map: Dict[str, str] = {}
    bibitem_pattern = r'\\bibitem\{([^}]+)\}(.*?)(?=\\bibitem|\s*\\end\{thebibliography\})'
    for match in re.finditer(bibitem_pattern, content, re.DOTALL):
        key = match.group(1).strip()
        raw_text = re.sub(r'[{}\\]', ' ', match.group(2)).strip()
        raw_text = re.sub(r'\s+', ' ', raw_text)
        bibitem_map[key] = raw_text[:300]

    citations = []
    seen_keys = set()

    # Step 2: For each \cite{key}, resolve through bibitem map
    cite_pattern = r'\\cite(?:\[[^\]]*\])?\{([^}]+)\}'
    for match in re.finditer(cite_pattern, content):
        keys = [k.strip() for k in match.group(1).split(',')]
        for key in keys:
            if key in seen_keys:
                continue
            seen_keys.add(key)
            raw = bibitem_map.get(key, '')  # use bibitem content if available
            citations.append({
                'key': key,
                'type': 'cite_resolved' if raw else 'cite_unresolved',
                'raw': raw if raw else key  # fall back to key if no bibitem found
            })

    # Step 3: Add any bibitem entries not referenced by \cite (standalone bibliography)
    for key, raw in bibitem_map.items():
        if key not in seen_keys:
            citations.append({
                'key': key,
                'type': 'bibitem',
                'raw': raw
            })

    return citations


def extract_citations_from_text(text: str) -> List[Dict[str, Any]]:
    """Extract citations from plain text using common patterns"""
    citations = []
    
    # Pattern 1: Numbered references [1], [2], etc.
    numbered_pattern = r'\[(\d+)\]'
    
    # Pattern 2: Author-year (Smith, 2020) or [Smith, 2020]
    author_year_pattern = r'\(([A-Z][a-z]+(?:\s+(?:and|&)\s+[A-Z][a-z]+)?,\s+\d{4})\)|\[([A-Z][a-z]+(?:\s+(?:and|&)\s+[A-Z][a-z]+)?,\s+\d{4})\]'
    
    # Pattern 3: DOI
    doi_pattern = r'doi:\s*(10\.\d{4,9}/[^\s]+)'
    for match in re.finditer(doi_pattern, text, re.IGNORECASE):
        citations.append({
            'key': match.group(1),
            'type': 'doi',
            'raw': match.group(0)
        })
    
    # Pattern 4: Bibliography section
    bib_section_pattern = r'(?:References|Bibliography|参考文献)[:\s]*\n(.*?)(?=\n\s*\n|$)'
    bib_match = re.search(bib_section_pattern, text, re.IGNORECASE | re.DOTALL)
    if bib_match:
        bib_text = bib_match.group(1)
        # Extract individual entries (numbered or bulleted)
        entry_pattern = r'^\s*(?:\[?\d+\]?\.?\s*|[-•]\s*)(.+)$'
        for line in bib_text.split('\n'):
            line = line.strip()
            if line:
                citations.append({
                    'key': line[:50],
                    'type': 'bibliography',
                    'raw': line
                })
    
    return citations


def extract_citations(file_path: str) -> List[Dict[str, Any]]:
    """Extract citations from a file based on its extension"""
    path = Path(file_path)
    
    if not path.exists():
        print(f"Error: File not found: {file_path}")
        return []
    
    suffix = path.suffix.lower()
    
    if suffix == '.docx':
        return extract_citations_from_docx(file_path)
    elif suffix in ['.md', '.txt', '.rst']:
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
        return extract_citations_from_text(text)
    elif suffix == '.tex':
        return extract_citations_from_tex(file_path)
    else:
        print(f"Warning: Unsupported file type: {suffix}")
        return []


def title_similarity(a: str, b: str) -> float:
    """Compute word-level Jaccard similarity between two titles"""
    if not a or not b:
        return 0.0
    stop = {'the', 'a', 'an', 'of', 'in', 'for', 'and', 'on', 'with', 'to', 'is', 'are', 'by'}
    def tokenize(s):
        words = re.findall(r'\b\w+\b', s.lower())
        return set(w for w in words if w not in stop and len(w) > 2)
    sa, sb = tokenize(a), tokenize(b)
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)


def verify_citation(citation: Dict[str, Any], threshold: float = 0.35) -> Dict[str, Any]:
    """Verify a citation against academic databases with title similarity check"""
    if not HAS_REQUESTS:
        return {
            'citation': citation,
            'status': 'error',
            'message': 'requests library not installed'
        }

    result = {
        'citation': citation,
        'status': 'unknown',
        'sources': []
    }

    raw_text = citation.get('raw', '')

    # --- Fast path: DOI verification ---
    doi_match = re.search(r'10\.\d{4,9}/[^\s,}]+', raw_text)
    if doi_match:
        doi = doi_match.group(0).rstrip('.')
        crossref_result = verify_doi_via_crossref(doi)
        if crossref_result:
            result['status'] = 'valid'
            result['sources'].append(crossref_result)
            return result

    # --- Title-based search ---
    query_title = extract_title_from_citation(raw_text)
    if not query_title:
        # If no title can be extracted and no DOI, mark as unverifiable (not hallucinated)
        result['status'] = 'unverifiable'
        return result

    # Search CrossRef
    time.sleep(0.3)  # polite delay
    crossref_results = search_crossref(query_title)
    for r in crossref_results:
        sim = title_similarity(query_title, r.get('title', ''))
        if sim >= threshold:
            result['status'] = 'valid'
            result['sources'].append({**r, 'similarity': round(sim, 3)})
            return result

    # Search Semantic Scholar
    time.sleep(0.3)
    semantic_results = search_semantic_scholar(query_title)
    for r in semantic_results:
        sim = title_similarity(query_title, r.get('title', ''))
        if sim >= threshold:
            result['status'] = 'valid'
            result['sources'].append({**r, 'similarity': round(sim, 3)})
            return result

    result['status'] = 'hallucinated'
    return result


def verify_doi_via_crossref(doi: str) -> Optional[Dict[str, Any]]:
    """Verify a DOI via CrossRef API"""
    try:
        url = f"https://api.crossref.org/works/{doi}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()['message']
            return {
                'source': 'CrossRef',
                'doi': doi,
                'title': data.get('title', [''])[0],
                'authors': [f"{a.get('given', '')} {a.get('family', '')}" for a in data.get('author', [])],
                'year': data.get('published', {}).get('date-parts', [['']])[0][0]
            }
    except Exception as e:
        pass
    return None


def search_crossref(query: str) -> List[Dict[str, Any]]:
    """Search CrossRef for a title"""
    try:
        url = "https://api.crossref.org/works"
        params = {'query.title': query, 'rows': 3}
        response = requests.get(url, params=params, timeout=10)
        if response.status_code == 200:
            results = []
            for item in response.json()['message']['items']:
                results.append({
                    'source': 'CrossRef',
                    'doi': item.get('DOI', ''),
                    'title': item.get('title', [''])[0],
                    'authors': [f"{a.get('given', '')} {a.get('family', '')}" for a in item.get('author', [])],
                    'year': item.get('published', {}).get('date-parts', [['']])[0][0]
                })
            return results
    except Exception as e:
        pass
    return []


def search_semantic_scholar(query: str) -> List[Dict[str, Any]]:
    """Search Semantic Scholar for a title"""
    try:
        url = "https://api.semanticscholar.org/graph/v1/paper/search"
        params = {'query': query, 'limit': 3, 'fields': 'title,authors,year,externalIds'}
        response = requests.get(url, params=params, timeout=10)
        if response.status_code == 200:
            results = []
            for item in response.json().get('data', []):
                results.append({
                    'source': 'Semantic Scholar',
                    'title': item.get('title', ''),
                    'authors': [a.get('name', '') for a in item.get('authors', [])],
                    'year': item.get('year', ''),
                    'paperId': item.get('paperId', '')
                })
            return results
    except Exception as e:
        pass
    return []


def extract_title_from_citation(text: str) -> Optional[str]:
    """Try to extract a paper title from citation text (supports LaTeX bibitem, plain text, Markdown)"""
    # LaTeX: ``Title'' or `Title'
    latex_quote = re.search(r"``([^']{10,}?)''", text)
    if latex_quote:
        return latex_quote.group(1).strip()

    # Straight double quotes: "Title"
    dquote = re.search(r'"([^"]{10,}?)"', text)
    if dquote:
        return dquote.group(1).strip()

    # Chinese title in 《》
    chinese_title = re.search(r'《([^》]+)》', text)
    if chinese_title:
        return chinese_title.group(1).strip()

    # "Title:" pattern
    title_label = re.search(r'(?:Title|题目)[:\s]+([^\n.]{10,})', text, re.IGNORECASE)
    if title_label:
        return title_label.group(1).strip()

    # Heuristic: after authors and year, extract remaining sentence as title
    # Matches: Author, A. (2020). Title sentence. Journal...
    heuristic = re.search(
        r'(?:[A-Z][a-z]+,?\s+(?:[A-Z]\.\s*)+)?(?:\(\d{4}\)\.?\s*)([A-Z][^.]{15,80})\.',
        text
    )
    if heuristic:
        return heuristic.group(1).strip()

    return None


def check_file(file_path: str, threshold: float = 0.85) -> Dict[str, Any]:
    """Check all citations in a file"""
    print(f"Checking citations in: {file_path}")
    
    # Extract citations
    citations = extract_citations(file_path)
    
    if not citations:
        print("No citations found in the file.")
        return {
            'file': file_path,
            'total_citations': 0,
            'valid': [],
            'suspicious': [],
            'hallucinated': [],
            'summary': {
                'valid_count': 0,
                'suspicious_count': 0,
                'hallucinated_count': 0
            }
        }
    
    print(f"Found {len(citations)} citation(s). Verifying...")

    # Verify each citation
    valid = []
    suspicious = []
    hallucinated = []
    unverifiable = []

    for citation in citations:
        result = verify_citation(citation, threshold)
        status = result['status']

        if status == 'valid':
            valid.append(result)
        elif status == 'suspicious':
            suspicious.append(result)
        elif status == 'unverifiable':
            unverifiable.append(result)
        else:
            hallucinated.append(result)

    # Build report
    report = {
        'file': file_path,
        'total_citations': len(citations),
        'valid': valid,
        'suspicious': suspicious,
        'hallucinated': hallucinated,
        'unverifiable': unverifiable,
        'summary': {
            'valid_count': len(valid),
            'suspicious_count': len(suspicious),
            'hallucinated_count': len(hallucinated),
            'unverifiable_count': len(unverifiable)
        }
    }

    return report


def print_report(report: Dict[str, Any]):
    """Print a human-readable report"""
    print("\n" + "="*60)
    print("CITATION VERIFICATION REPORT")
    print("="*60)
    print(f"File: {report['file']}")
    print(f"Total citations: {report['total_citations']}")
    print(f"  ✅ Valid:         {report['summary']['valid_count']}")
    print(f"  ⚠️  Suspicious:    {report['summary']['suspicious_count']}")
    print(f"  ❌ Hallucinated:  {report['summary']['hallucinated_count']}")
    print(f"  ❓ Unverifiable:  {report['summary'].get('unverifiable_count', 0)}")

    if report['hallucinated']:
        print("\n❌ HALLUCINATED CITATIONS (not found in any database):")
        for item in report['hallucinated']:
            raw = item['citation'].get('raw', item['citation'].get('key', 'Unknown'))
            print(f"  - [{item['citation']['key']}] {raw[:120]}")

    if report['suspicious']:
        print("\n⚠️  SUSPICIOUS CITATIONS:")
        for item in report['suspicious']:
            print(f"  - {item['citation'].get('raw', item['citation'].get('key', 'Unknown'))[:120]}")

    if report.get('unverifiable'):
        print("\n❓ UNVERIFIABLE (no title could be extracted):")
        for item in report['unverifiable']:
            print(f"  - [{item['citation']['key']}]")

    print("="*60 + "\n")


def main():
    parser = argparse.ArgumentParser(description='Check citations in documents')
    parser.add_argument('--file', type=str, help='Single file to check')
    parser.add_argument('--dir', type=str, help='Directory to scan for documents')
    parser.add_argument('--output', type=str, help='Save report to JSON file')
    parser.add_argument('--threshold', type=float, default=0.85, help='Similarity threshold (default: 0.85)')
    
    args = parser.parse_args()
    
    if not args.file and not args.dir:
        parser.print_help()
        sys.exit(1)
    
    reports = []
    
    if args.file:
        report = check_file(args.file, args.threshold)
        print_report(report)
        reports.append(report)
    
    if args.dir:
        dir_path = Path(args.dir)
        for ext in ['*.docx', '*.md', '*.txt', '*.tex']:
            for file_path in dir_path.glob(f'**/{ext}'):
                report = check_file(str(file_path), args.threshold)
                print_report(report)
                reports.append(report)
    
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(reports, f, indent=2, ensure_ascii=False)
        print(f"\nReport saved to: {args.output}")


if __name__ == '__main__':
    main()
