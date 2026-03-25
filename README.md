# md-to-docx

A Claude Code skill for converting Markdown files to professionally formatted Word (.docx) documents with automatic template style detection and Mermaid diagram rendering.

## Features

- **Template-aware formatting**: Automatically extracts fonts, sizes, alignment, spacing, and margins from any `.doc` or `.docx` template
- **Chinese typesetting defaults**: 宋体/黑体, 五号字 (10.5pt), 首行缩进两字符 when no template is provided
- **Mermaid diagram rendering**: Converts Mermaid code blocks to PNG images via `mmdc` and embeds them in the document
- **Full Markdown support**: Headings (h1-h6), bold/italic/code/strikethrough, bullet/ordered lists, tables, block quotes, code blocks
- **Cross-platform**: Works on macOS, Windows, and Linux with automatic Chrome detection

## Installation

### As a Claude Code Skill

1. Clone this repo:
   ```bash
   git clone https://github.com/xielaoban/md-to-docx.git
   ```

2. Add to your Claude Code settings:
   ```bash
   claude skill add /path/to/md-to-docx
   ```

### Standalone Script

```bash
pip install python-docx
npm install -g @mermaid-js/mermaid-cli
```

## Usage

### Via Claude Code

Just ask naturally:
- "把 chapter01.md 转成 Word"
- "Convert this markdown to docx using my template"
- "Export the report to Word format"

### Via Command Line

```bash
# Basic conversion (default Chinese typesetting)
python3 scripts/md_to_docx.py input.md -o output.docx

# With template (auto-detect formatting)
python3 scripts/md_to_docx.py input.md -o output.docx -t template.doc

# With .docx template
python3 scripts/md_to_docx.py input.md -o output.docx -t template.docx
```

## Dependencies

| Tool | Purpose | Install |
|------|---------|---------|
| `python-docx` | Word document generation | `pip install python-docx` |
| `mmdc` | Mermaid diagram rendering | `npm install -g @mermaid-js/mermaid-cli` |
| Chrome/Chromium | Mermaid rendering backend | Pre-installed on most systems |

## Template Format Detection

When a template is provided, the script analyzes it to identify:

1. **Title hierarchy** — identified by font size and alignment (largest centered = chapter title)
2. **Body text** — identified by paragraphs with first-line indentation
3. **Page setup** — margins, paper size copied from template
4. **Chinese font pairing** — preserved (e.g., Times New Roman ↔ 宋体)

### Default Formatting (no template)

| Element | Font | Size | Alignment |
|---------|------|------|-----------|
| Chapter title (h1) | 黑体 | 18pt | Center |
| Section title (h2) | 黑体 | 16pt | Center |
| Subsection (h3) | 黑体 | 14pt | Left |
| Body text | 宋体 | 10.5pt | Justify, 2-char indent |
| Code | Consolas | 9pt | Left |
| Quote | 楷体 | 10.5pt | Italic, indented |

## Supported Markdown

- Headings: h1–h6
- Inline formatting: **bold**, *italic*, `code`, ~~strikethrough~~, ***bold italic***
- Lists: bullet (`-`/`*`) and ordered (`1.`)
- Tables with header row
- Code blocks with language label
- Mermaid diagrams (rendered to PNG)
- Block quotes and horizontal rules

## License

MIT
