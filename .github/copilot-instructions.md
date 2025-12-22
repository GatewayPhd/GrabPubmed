# Copilot Instructions for PubMed Info Grabber

## Project Overview
This is a PubMed literature retrieval and analysis tool that queries NCBI's PubMed API, extracts article metadata, enriches it with journal impact factors, and generates styled interactive HTML reading lists with sidebar navigation and persistent state management. The workflow is Jupyter notebook-driven with modular Python utilities.

## Architecture & Data Flow

**Main Components:**
- [pubmed_utils.py](pubmed_utils.py): Core `pubmed_utils` class with 6 methods for PubMed API interaction and data processing
- [html_generate.py](html_generate.py): HTML generation with interactive features (sidebar, star/read buttons, localStorage persistence)
- [pumbed_query.ipynb](pumbed_query.ipynb): Main workflow orchestration notebook with comprehensive markdown documentation
- [html_generate_v2.py](html_generate_v2.py): Alternative/backup HTML generator (use [html_generate.py](html_generate.py) for production)
- [test_updates.py](test_updates.py): Development testing script (not part of main workflow)

**Typical Workflow (executed cell-by-cell):**
1. Configure search parameters (API key, keywords, date range, paper type, limits)
2. Call `get_main_info_into_excel()` → queries PubMed, saves PMID/Title/Journal/Abstract/DOI to Excel
3. Call `embed_IF_into_excel()` → scrapes journal impact factors from ScienceDirect, adds IF/Quartile columns
4. Call `generate_reading_list()` → reads Excel, generates interactive HTML with sidebar bookmarks and persistent state

**Output Structure:**
```
paper_donload/
├── {query_name}.xlsx              # Structured metadata table (11 columns)
└── {query_name}_reading_list.html # Interactive reading list with sidebar navigation
```

## Critical Patterns

### Module Reloading Protocol
When modifying [html_generate.py](html_generate.py) or [pubmed_utils.py](pubmed_utils.py), **always reload before use** in notebook cells:
```python
import importlib
import html_generate
importlib.reload(html_generate)
from html_generate import generate_reading_list
```
This is necessary because Jupyter caches imports. Check [pumbed_query.ipynb](pumbed_query.ipynb) cells for examples.

### PubMed Query Syntax
Uses NCBI E-utilities advanced search syntax (NOT standard boolean):
- Operators: `AND`, `OR`, `NOT` (uppercase, space-separated)
- Field tags: `[Title]`, `[Title/Abstract]`, `[Author]`, `[Journal]`, `[MeSH Terms]`, `[Affiliation]`
- Wildcards: `fibro*` matches "fibroblast", "fibrosis", etc.
- Grouping: `(wnt5a OR wnt7a) AND cancer`

Paper type filter is added programmatically via `[PT]` tag (e.g., `"Journal Article"[PT]`).

### Excel Column Schema
Fixed 11-column structure (see `excel_property_dic` in [pubmed_utils.py](pubmed_utils.py)):
1. PMID, 2. Title, 3. Journal, 4. IF, 5. JCR_Quartile, 6. CSA_Quartile, 7. Top, 8. Open Access (OA), 9. publish_date, 10. Abstract, 11. DOI

**Critical:** Do NOT reorder columns—HTML generation expects this sequence. Column names changed from original format `'Title (TI)'` to simple `'Title'` for better compatibility.

### IF Scraping Logic
`embed_IF_into_excel()` performs fuzzy journal name matching:
1. First tries exact match against ScienceDirect table
2. Falls back to fuzzy matching with `fuzzy_match_score()` (>60% threshold)
3. Uses `refine_IF_matching()` for manual override/correction

**Common issue:** Journal name variations ("J. Biol. Chem." vs "Journal of Biological Chemistry") may cause mismatches. Check Excel IF column for empty values.

### HTML Interactive Features
[html_generate.py](html_generate.py) includes three interactive features with localStorage persistence:

1. **Sidebar Bookmark Navigation** (☰):
   - Fixed position sidebar (280px wide) with collapsible panel
   - Bookmark format: `{Journal}. {YYYYMMDD}` (e.g., "Nat Commun. 20251216")
   - Each bookmark has `data-article-id` attribute and `<span id="indicators-{idx}">` for status icons
   - Toggle button (☰) at top-left to show/hide sidebar
   - Body padding adjusts dynamically: `padding-left: 300px` (sidebar shown) or `0` (hidden)
   - Click any bookmark to smooth-scroll to corresponding article
   
2. **Star Function** (⭐): Mark important papers
   - Starred cards show gold left border (4px solid #ffd700)
   - Click star button to toggle on/off
   - State persisted in localStorage key 'starred'
   - Sidebar shows ⭐ icon for starred articles
   
3. **Read Function** (✓): Mark papers as read
   - Read cards reduce opacity to 0.6
   - Helps track reading progress
   - State persisted in localStorage key 'read'
   - Sidebar shows ✓ icon (green #4CAF50) for read articles

**State Synchronization Flow:**
1. User clicks ⭐ or ✓ button in article card
2. `toggleStar(this)` or `toggleRead(this)` receives button element as parameter (NOT article ID)
3. Function finds parent `.card` via `btn.closest('.card')`, gets `card.id`
4. Updates localStorage with array of article IDs
5. Calls `updateSidebarIndicator(articleId)` to update corresponding bookmark
6. Sidebar bookmark's `<span id="indicators-{idx}">` dynamically populates with ⭐ and/or ✓

**Critical:** Button `onclick` attributes MUST use `onclick="toggleStar(this)"` not `onclick="toggleStar('{article_id}')"`. Functions expect DOM element reference.

### HTML Keyword Highlighting
`_build_pattern_from_query()` in [html_generate.py](html_generate.py):
- Parses search query, removes `AND/OR/NOT` operators
- Ignores field tags (e.g., `[Title]`)
- Converts wildcards (`fibro*` → regex `\w*`)
- Skips terms after `NOT` to avoid highlighting excluded words

## Development Workflows

### Testing New Queries
Modify parameters in [pumbed_query.ipynb](pumbed_query.ipynb) cell 2:
- Set `grab_total = 20` for quick tests (instead of `None`)
- Use recent `release_date_cutoff` (e.g., 90 days) to limit results
- Always check `print(f"Find total: {total}")` output before fetching

### Adding New Excel Columns
1. Update `excel_property_dic` dictionary in `__init__` or `get_main_info_into_excel()`
2. Add header in row 1: `ws.cell(row=1, column=...).value = "New Column"`
3. Populate data in fetch loop using `ws.cell(row=cur_row, column=...)`
4. Update [html_generate.py](html_generate.py) if column affects rendering

### Debugging HTML Output
- HTML uses inline CSS (no external stylesheets)
- Keyword highlighting: check `pattern` variable in `generate_reading_list()`
- Search info section: controlled by `search_info` dict with keys: `search_keywords`, `paper_type`, `release_date_cutoff`, `grab_total`, `save_path`, `search_date`
- JavaScript errors: Check browser console (F12) for syntax errors
- Sidebar not showing bookmarks: Verify Excel has `Journal` and `publish_date` columns populated
- Buttons not working: Ensure onclick uses `this` parameter, check JavaScript function signatures

### Modifying Interactive Features
When updating sidebar/button functionality:
1. **HTML Structure**: Sidebar links need `data-article-id="{idx}"` and `<span id="indicators-{idx}"></span>`
2. **CSS Styles**: `.bookmark-indicators { display: inline-flex; gap: 3px; }`, `.star-indicator { color: #ffd700; }`, `.read-indicator { color: #4CAF50; }`
3. **JavaScript Functions**: 
   - `toggleStar(btn)` and `toggleRead(btn)` receive button DOM element
   - `updateSidebarIndicator(articleId)` updates single bookmark
   - `updateAllSidebarIndicators()` initializes all bookmarks on page load
   - `window.onload` must call `updateAllSidebarIndicators()`
4. **State Management**: Use `localStorage.getItem('starred')` and `localStorage.getItem('read')` for persistence

## Dependencies
Core packages (install via pip):
- `biopython` (Bio.Entrez for PubMed API)
- `pandas` (Excel reading)
- `openpyxl` (Excel writing)
- `requests` (HTTP for IF scraping)
- `beautifulsoup4` (HTML parsing for IF tables)
- `tqdm` (progress bars)

## Common Issues

**"NCBI API rate limit exceeded"**: Get free API key from https://www.ncbi.nlm.nih.gov/account/ (increases limit from 3 to 10 req/sec)

**Empty IF column**: Journal name mismatch—check ScienceDirect uses different abbreviation. Use `refine_IF_matching()` to manually map.

**HTML not updating**: Forgot `importlib.reload(html_generate)` before calling `generate_reading_list()`.

**Excel overwrite warning**: `get_main_info_into_excel()` creates new workbook—existing file will be replaced. Save backups if needed.

**Independent IF update**: `embed_IF_into_excel(excel_path)` can be called standalone to update IF information in existing Excel files without re-querying PubMed.

**HTML buttons not clickable**: Check button onclick attributes use `onclick="toggleStar(this)"` format, not `onclick="toggleStar('{article_id}')"`. JavaScript functions need DOM element reference.

**Sidebar bookmarks show "Unknown"**: Excel column names must be `'Journal'` and `'publish_date'` (not `'Journal (TA)'` or `'Publish Date (LR)'`). Sidebar generation uses these exact names with fallback logic.

**JavaScript syntax errors**: Verify brace balance in `<script>` section. Use `content.count('{') == content.count('}')` to check. Common issue: duplicate function definitions or unclosed blocks.

**Sidebar not syncing with star/read**: Ensure `updateSidebarIndicator(articleId)` is called in both `toggleStar()` and `toggleRead()` functions after `localStorage.setItem()`.

## Project Structure (Updated for v2.0)

```
GrabPubmed/
├── .github/
│   └── copilot-instructions.md     # This file - development guidelines
├── paper_donload/                  # Output directory
│   ├── .gitkeep                    # Track empty directory in git
│   ├── *.xlsx                      # Excel files with metadata
│   └── *_reading_list.html         # Generated HTML reading lists
├── pumbed_query.ipynb              # Main workflow notebook (with markdown documentation)
├── pubmed_utils.py                 # PubMed API interaction & IF scraping
├── html_generate.py                # HTML generation with interactive features (PRODUCTION)
├── html_generate_v2.py             # Alternative HTML generator (backup/experimental)
├── test_updates.py                 # Development testing script
├── README.md                       # User-facing documentation
├── requirements.txt                # Python dependencies
├── LICENSE                         # MIT License
└── .gitignore                      # Git ignore patterns
```

## Version History

**v2.0 (2025-12-22):**
- Added collapsible sidebar navigation with bookmark links
- Implemented star and read marking with persistent state
- Added real-time status synchronization between article cards and sidebar bookmarks
- Optimized for GitHub with comprehensive documentation
- Simplified Excel column names for better compatibility

**v1.0 (Original):**
- PubMed query and metadata extraction
- Impact factor scraping from ScienceDirect
- Basic HTML generation with keyword highlighting

## Contributing Guidelines

When making changes:
1. Test with a small dataset (`grab_total = 20`) first
2. Always reload modules in notebook after code changes
3. Verify HTML output in browser (Chrome/Firefox/Edge)
4. Check browser console for JavaScript errors
5. Update documentation (README.md and this file) for new features
6. Ensure backward compatibility with existing Excel files
