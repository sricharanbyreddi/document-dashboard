# Document Dashboard (Python / Flask)

A minimal local document viewer based on the supplied screenshot: numbered file list on the left, document viewer on the right, Add documents button, and dark/light theme toggle.

## Supported formats
- Word: `.docx`
- Markdown: `.md`, `.markdown`
- Code/text: `.py`, `.sql`, `.js`, `.ts`, `.java`, `.json`, `.yaml`, `.yml`, `.html`, `.css`, `.xml`, `.csv`, `.log`, `.txt`

## Setup on Windows
1. Extract this ZIP.
2. Open the extracted folder in VS Code.
3. Open Terminal > New Terminal and run:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

4. Visit http://127.0.0.1:5000

You can also double-click `run.bat`.

## Rendering improvements
- Markdown source stored inside Word documents is rendered as Markdown/HTML, including fenced SQL and other code blocks, Markdown tables, and Mermaid diagrams.
- Embedded DOCX images are represented as inline data URLs and placed where their image run occurs in the document instead of being collected at the end.
- Code blocks preserve indentation and can scroll horizontally for long SQL statements.

All processing runs locally. Mermaid diagrams use the Mermaid browser library from a CDN, so diagram rendering requires an internet connection; the rest of the dashboard works locally.
