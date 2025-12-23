# PubMed Literature Retrieval and Interactive Reading List Generator

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Jupyter](https://img.shields.io/badge/Jupyter-Notebook-orange.svg)](https://jupyter.org/)

An automated tool for querying PubMed, extracting article metadata, enriching with journal impact factors, and generating interactive HTML reading lists with sidebar navigation and persistent state management.

## ✨ Features

### Core Functionality
- 🔍 **Advanced PubMed Search**: Full support for E-utilities query syntax with field tags, boolean operators, and wildcards
- 📊 **Impact Factor Integration**: Automatic scraping of IF and Quartile information from ScienceDirect
- 📁 **Structured Export**: Saves metadata to Excel with 11 columns (PMID, Title, Journal, IF, Quartile, Abstract, DOI, etc.)
- 🌐 **Interactive HTML**: Beautiful night-mode reading list with full interactivity

### HTML Reading List Features
- 🌙 **Night Mode Design**: Dark gradient background optimized for comfortable reading
- 🎨 **Keyword Highlighting**: Automatic highlighting of search terms in titles (yellow) and abstracts (orange)
- 📑 **Collapsible Sidebar**: 
  - Navigate between articles with `Journal. YYYYMMDD` bookmarks
  - Real-time status indicators: ⭐ (starred), ✓ (read)
  - Smooth show/hide transitions
- ⭐ **Star System**: Mark important papers with persistent state
- ✓ **Read Tracking**: Track reading progress across sessions
- 💾 **Persistent State**: All user interactions saved in browser localStorage

## 🚀 Quick Start

### Prerequisites

```bash
pip install biopython pandas openpyxl requests beautifulsoup4 tqdm
```

### Basic Usage

1. **Clone the repository**:
```bash
git clone https://github.com/yourusername/grab-pubmed-info.git
cd grab-pubmed-info
```

2. **Open the notebook**:
```bash
jupyter notebook pumbed_query.ipynb
```

3. **Configure your search** (Cell 2):
```python
api_key = "your_ncbi_api_key"  # Get from https://www.ncbi.nlm.nih.gov/account/
search_key_words = "(wnt5a NOT cancer) AND fibro*"
release_date_cutoff = 365  # Papers from last year
paper_type = "Journal Article"
save_path = "./paper_donload/my_query.xlsx"
```

4. **Run all cells** to:
   - Query PubMed
   - Fetch metadata
   - Scrape impact factors
   - Generate interactive HTML

5. **Open the HTML file** in your browser to start reading!

## 📚 Documentation

### PubMed Query Syntax

The tool supports full NCBI E-utilities advanced search syntax:

**Boolean Operators:**
```python
search_key_words = "wnt5a AND cancer"        # Both terms
search_key_words = "wnt5a OR wnt7a"          # Either term
search_key_words = "wnt5a NOT cancer"        # Exclude term
search_key_words = "(wnt5a OR wnt7a) AND cancer"  # Combined
```

**Field Tags:**
```python
search_key_words = "wnt5a[Title]"                        # Title only
search_key_words = "wnt5a[Title/Abstract]"               # Title or Abstract
search_key_words = "Smith J[Author]"                     # Specific author
search_key_words = "Nature[Journal]"                     # Specific journal
search_key_words = "breast cancer AND China[Affiliation]"  # Institution
```

**Wildcards:**
```python
search_key_words = "fibro*"  # Matches: fibroblast, fibrosis, fibrotic, etc.
```

### Excel Column Schema

Generated Excel files have 11 columns:

| Column | Description |
|--------|-------------|
| PMID | PubMed unique identifier |
| Title | Article title |
| Journal | Journal abbreviation |
| IF | Impact Factor (from ScienceDirect) |
| JCR_Quartile | JCR Quartile (Q1/Q2/Q3/Q4) |
| CSA_Quartile | CSA Quartile |
| Top | Top journal indicator |
| Open Access | OA status |
| publish_date | Publication date (YYYYMMDD) |
| Abstract | Full abstract text |
| DOI | Digital Object Identifier |

### HTML Interface Guide

**Sidebar Navigation:**
- Click `☰` button to toggle sidebar
- Bookmarks format: `Nat Commun. 20251216`
- ⭐ = Starred articles
- ✓ = Read articles

**Article Cards:**
- Click ⭐ to star important papers (gold left border appears)
- Click ✓ to mark as read (card opacity reduces to 0.6)
- All states persist across browser sessions

## 🛠️ Advanced Usage

### Independent IF Update

Update impact factors for existing Excel files without re-querying PubMed:

```python
from pubmed_utils import pubmed_utils

utils = pubmed_utils()
utils.embed_IF_into_excel('./paper_donload/existing_file.xlsx')
```

### Batch Processing

Process multiple queries:

```python
queries = [
    ("wnt5a AND fibrosis", "wnt5a_fibrosis.xlsx"),
    ("wnt7a AND regeneration", "wnt7a_regen.xlsx"),
]

for keywords, path in queries:
    utils.get_main_info_into_excel(api_key, keywords, 365, "Journal Article", None, path)
    utils.embed_IF_into_excel(path)
    generate_reading_list(path, path.replace('.xlsx', '_reading_list.html'))
```

### Custom HTML Styling

Modify `html_generate.py` to customize:
- Colors (CSS variables in `<style>` section)
- Layout (adjust `.card`, `.sidebar` styles)
- Highlighting patterns (`_build_pattern_from_query()` function)

## 📊 Project Structure

```
grab-pubmed-info-master/
├── pumbed_query.ipynb          # Main workflow notebook (⭐ Start here)
├── pubmed_utils.py             # PubMed API & IF scraping logic
├── html_generate.py            # HTML generation with interactivity
├── paper_donload/              # Output directory (auto-created)
│   ├── *.xlsx                  # Excel files with metadata
│   └── *_reading_list.html     # Interactive HTML reading lists
├── README.md                   # This file
├── LICENSE                     # MIT License
└── requirements.txt            # Python dependencies
```

## 🐛 Troubleshooting

### Common Issues

**Problem:** `NCBI API rate limit exceeded`  
**Solution:** Get free API key from https://www.ncbi.nlm.nih.gov/account/ (increases limit from 3 to 10 req/sec)

**Problem:** Empty IF column in Excel  
**Solution:** Journal name mismatch - use `refine_IF_matching()` method for manual correction

**Problem:** HTML buttons not clickable  
**Solution:** Ensure you're using a modern browser (Chrome/Firefox/Edge). Check browser console for JavaScript errors.

**Problem:** Sidebar bookmarks show "Unknown"  
**Solution:** Verify Excel has `Journal` and `publish_date` columns properly populated

**Problem:** HTML not updating after code changes  
**Solution:** Use `importlib.reload(html_generate)` before calling `generate_reading_list()`

### Error Reporting

Found a bug? Please [open an issue](https://github.com/yourusername/grab-pubmed-info/issues) with:
- Error message
- Python version
- Browser (for HTML issues)
- Minimal reproducible example

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Development Setup

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 Citation

If you use this tool in your research, please cite:

```bibtex
@software{pubmed_info_grabber2025,
  author = {Li, Xiang and GitHub Copilot},
  title = {PubMed Literature Retrieval and Interactive Reading List Generator},
  year = {2025},
  url = {https://github.com/GatewayPhd/grab-pubmed-info}
}
```

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👥 Authors

- **李想 (Li Xiang)** - Initial work and concept
- **GitHub Copilot** - Interactive HTML features, code optimization

## 🙏 Acknowledgments

- NCBI for providing the E-utilities API
- ScienceDirect for impact factor data
- The Python scientific computing community

