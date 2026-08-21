# Real-Time MCQ Extractor, Preview & Spell Checker

A stateless web application for extracting multiple-choice questions from table-formatted `.docx` files, rendering a live browser preview, detecting repeated questions, preserving Word equations and embedded images, and highlighting spelling issues.

Built with **FastAPI**, **python-docx**, **MathJax**, and a lightweight HTML/CSS/JavaScript frontend.

## Key Features

- Upload `.docx` files directly from the browser
- Parse MCQs from Word table cells using fixed cell indexes
- Render a clean real-time preview without saving data to a database
- Preserve DOCX Word equations by converting OMML equations to MathML
- Display embedded DOCX images using in-memory data URIs
- Show options in a two-column MCQ layout
- Highlight the correct answer in a separate answer block
- Render explanations/solutions in a dedicated explanation block
- Detect repeated questions in memory
- Show duplicate warning banners with jump-scroll links
- Check English spelling through LanguageTool
- Check Bangla spelling through a configurable local wordlist
- Docker-ready and suitable for deployment on Render, VPS, or similar platforms

## Architecture

This app is intentionally stateless.

- Uploaded files are processed in memory
- Parsed MCQ data is returned directly to the browser
- No database is used
- No question data is stored by the application

## Input DOCX Format

Each MCQ should be stored as one table with at least 8 rows.

| Cell | Data |
| --- | --- |
| `[0,0]` | Serial number |
| `[0,1]` | Category |
| `[1,0]` | Main question |
| `[2,0]` | Option A |
| `[3,0]` | Option B |
| `[4,0]` | Option C |
| `[5,0]` | Option D |
| `[6,0]` | Explanation / solve |
| `[7,0]` | Correct answer key: `A`, `B`, `C`, `D`, or Bangla option letters |

## Local Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open:

```text
http://127.0.0.1:8000
```

If port `8000` is already in use:

```powershell
uvicorn app.main:app --reload --port 8001
```

## Spell Checking

### English

English spell checking is handled through the LanguageTool API.

Default endpoint:

```text
https://api.languagetool.org/v2/check
```

### Bangla

Bangla spell checking is handled locally with a wordlist.

Default wordlist:

```text
data/bn_words.txt
```

You can replace or extend this file with a larger curated Bangla dictionary for better accuracy.

## Environment Variables

```powershell
$env:LANGUAGETOOL_URL = "https://api.languagetool.org/v2/check"
$env:SPELLCHECK_ENABLED = "true"
$env:MAX_UPLOAD_BYTES = "26214400"
$env:BN_WORDLIST_PATH = "C:\path\to\bn_words.txt"
```

## Equations And Images

DOCX equations are read from Word OMML XML and converted to MathML. MathJax renders the MathML in the browser.

Embedded images are extracted from the DOCX package and returned as in-memory data URIs. Image dimensions from Word are preserved when available.

## API Endpoints

| Method | Endpoint | Purpose |
| --- | --- | --- |
| `GET` | `/` | Web interface |
| `POST` | `/api/parse` | Upload and parse a `.docx` file |
| `POST` | `/api/spellcheck` | Spell-check parsed text segments |

## Docker

```powershell
docker build -t mcq-proofreader .
docker run --rm -p 8000:8000 mcq-proofreader
```

Open:

```text
http://127.0.0.1:8000
```

## Project Structure

```text
MCQ_Proofreader/
|-- app/
|   |-- main.py
|   `-- parser.py
|-- data/
|   `-- bn_words.txt
|-- static/
|   |-- app.js
|   `-- styles.css
|-- templates/
|   `-- index.html
|-- tests/
|   `-- test_parser.py
|-- Dockerfile
|-- requirements.txt
`-- README.md
```

## Notes

- The app does not store uploaded files.
- Bangla spell-check quality depends on the provided wordlist.
- LanguageTool availability controls English spell-check results.
- Very complex Word equations may require additional OMML-to-MathML mapping rules.

## Credit

Powered by [Mazharul](http://mazharul.dev.cv/).
