# CEFR Vocab Analyzer

A Python tool that analyzes the difficulty of English texts based on the CEFR (Common European Framework of Reference) vocabulary grading system (A1–C2).

## Motivation

When self-studying English, it's hard to know whether a book or article is at the right difficulty level. This tool solves that problem by automatically analyzing every word in a text and producing a detailed difficulty report.

## Features

- 📊 Analyzes `.txt` and `.epub` files
- 🔤 Lemmatizes words to their base form before lookup
- 🏷️ Grades each word by CEFR level (A1–C2) and calculates a weighted difficulty score
- 📁 Outputs an Excel report with:
  - **Sheet 1**: Overall difficulty summary and CEFR breakdown chart
  - **Sheets 2–7**: Word lists per CEFR level with dictionary definitions
  - **Last sheet**: Uncategorized words (slang, neologisms, proper nouns)
- 🧠 Maintains a personal vocabulary bank (`is_known.xlsx`):
  - Mark words as "learned" via dropdown in the report
  - Learned words (with your notes) are automatically carried over to the vocab bank
  - On the next run, learned words are filtered out of the report

## How It Works

```
Input (.txt / .epub)
      ↓
Tokenize & Lemmatize (NLTK)
      ↓
Look up local dictionary first (ECDICT) → API fallback if not found
      ↓
Grade by CEFR level
      ↓
Output Excel report + update is_known.xlsx
```

## Requirements

- Python 3.x
- Required packages:

| Package | Purpose |
|---------|---------|
| `nltk` | Tokenization and lemmatization |
| `cefrpy` | CEFR level grading per word |
| `ebooklib` | Reading `.epub` files |
| `beautifulsoup4` | Parsing HTML content inside epub |
| `pandas` | Data processing |
| `openpyxl` | Excel report generation |
| `requests` | API fallback for dictionary lookup |

Install dependencies:
```bash
pip install nltk cefrpy ebooklib beautifulsoup4 pandas openpyxl requests
```

## Setup

1. Clone this repository:
```bash
git clone https://github.com/claire77chang-sudo/CEFR-vocab-analyzer.git
cd CEFR-vocab-analyzer
```

2. Download the ECDICT local dictionary CSV from:
👉 https://github.com/skywind3000/ECDICT

   Place `ecdict.csv` in the same folder as `CEFR_vocab_analyzer.py`.

3. Place the `.txt` or `.epub` file you want to analyze in the same folder.

4. Run the script:
```bash
python CEFR_vocab_analyzer.py
```

## Output Files

| File | Description |
|------|-------------|
| `[book_name]_vocabulary_report.xlsx` | Difficulty report for the analyzed text |
| `is_known.xlsx` | Personal vocabulary bank (auto-created on first run) |

## Known Limitations

- PDF files are not yet supported
- Word sense disambiguation is not implemented (words are matched by base form only, not by context)
- No graphical user interface — runs entirely in the terminal
- Processing time may take several minutes for longer texts

## Future Plans

- [ ] Add PDF support
- [ ] Context-aware word sense disambiguation
- [ ] Graphical user interface (GUI)
- [ ] Faster lookup performance

## License

This project is for personal and educational use.
