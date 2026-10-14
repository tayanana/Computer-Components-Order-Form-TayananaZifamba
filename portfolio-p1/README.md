# Portfolio Part 1 – Conception Phase (DLBSEPCP01_E, Task 1)

| File | Purpose |
| --- | --- |
| `Zifamba-Tayanana_3220254_Cloud Programming_P1_S.html` | Source document: one A4 page, inline SVG architecture diagram |
| `Zifamba-Tayanana_3220254_Cloud Programming_P1_S.pdf` | Submission file for PebblePad (printed from the HTML) |
| `Zifamba-Tayanana_3220254_Cloud Programming_P1_S.docx` | Editable Word version (same text, diagram embedded as PNG) |
| `architecture-diagram.png` | 300 dpi diagram export (PNG required by the assignment) |
| `build_docx.py` | Regenerates the PNG and DOCX from the HTML/PDF |

Regenerate after editing the HTML:

```bash
google-chrome --headless=new --no-sandbox --no-pdf-header-footer \
  --print-to-pdf="Zifamba-Tayanana_3220254_Cloud Programming_P1_S.pdf" \
  "file://$PWD/Zifamba-Tayanana_3220254_Cloud Programming_P1_S.html"
pip install python-docx pymupdf lxml && python3 build_docx.py
```
