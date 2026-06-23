# Word Live MCP Server

A Windows MCP (Model Context Protocol) server that attaches to a **running Microsoft Word instance** via COM automation and exposes 20 tools for reading, editing, formatting, and exporting documents — live on screen.

The Word counterpart to the [PowerPoint Live MCP](https://github.com/kennypeh85/powerpoint-mcp).

## How It Works

The server attaches to the running `Word.Application` instance via `win32com` (pywin32). Because it talks to the *live COM object*, every edit appears instantly on the Word canvas — no file round-trip, no reopen.

## Tools (20)

| Category | Tool | Description |
|----------|------|-------------|
| **Discovery** | `list_documents` | List all open Word documents |
| | `get_active_document` | Details of active document (name, path, word/para count) |
| **Read** | `get_document_content` | Paginated paragraph listing with styles |
| | `get_paragraph` | Full text + formatting of one paragraph |
| | `get_table` | All cell content from a table |
| | `get_selection` | Current cursor/selection state |
| | `find_text` | Full-text search with paragraph locations |
| **Edit** | `set_paragraph_text` | Replace text of a paragraph |
| | `insert_paragraph_after` | Insert new paragraph after a given index |
| | `delete_paragraph` | Delete a paragraph |
| | `replace_text` | Find-and-replace throughout document |
| **Format** | `set_font` | Font size/bold/italic/color/name |
| | `set_highlight` | Yellow/red/green/none highlight |
| | `set_paragraph_style` | Apply named style (Heading 1, Normal, etc.) |
| **Tables** | `set_table_cell` | Set cell text by (row, column) |
| | `insert_table` | Insert a new table |
| **Export** | `export_pdf` | Export document to PDF |
| | `export_page_image` | Export single page as PNG (via PDF render) |
| **Other** | `update_toc` | Update all Tables of Contents |
| | `save_document` | Save the document |

## Requirements

- **Windows** + Microsoft Word installed
- **Python 3.11+** with these packages:
  ```bash
  pip install pywin32 mcp pymupdf
  ```
  Run `python Scripts/pywin32_postinstall.py -install` once after installing pywin32.

## Installation

### Hermes Agent

Add to `~/.hermes/config.yaml`:
```yaml
mcp_servers:
  word:
    command: "C:/Users/You/AppData/Local/Programs/Python/Python313/python.exe"
    args: ["C:/path/to/word-live-mcp/server.py"]
    connect_timeout: 30
    enabled: true
    timeout: 90
```

### Claude Desktop

Add to `claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["C:/path/to/word-live-mcp/server.py"]
    }
  }
}
```

### Generic MCP Client

The server communicates over **stdio** using the MCP protocol. Any MCP-compatible client can connect by spawning the server process.

## Usage

**Word must be running** with a document open. Tools operate on the active document by default; pass `doc="filename.docx"` to target a specific open document.

Paragraph and table indices are **1-indexed** (matching what you see in Word).

### Example workflows

- **"Read me paragraph 5"** → `get_paragraph(paragraph_index=5)`
- **"Find all mentions of 'confidential'"** → `find_text(find="confidential")`
- **"Replace 'old term' with 'new term'"** → `replace_text(find="old term", replace="new term")`
- **"Highlight paragraph 20 in yellow"** → `set_highlight(paragraph_index=20, color="yellow")`
- **"Show me what page 3 looks like"** → `export_page_image(page=3)`
- **"Set paragraph 10 font to bold red"** → `set_font(paragraph_index=10, bold=True, color_rgb="FF0000")`

## Architecture

- **COM session per call**: Each tool call re-initialises COM (`pythoncom.CoInitialize`) and dispatches `Word.Application`, which transparently reconnects to the running instance.
- **`app` parameter injection**: The `word_tool` decorator injects the COM app handle as the first argument, but strips it from the public signature so FastMCP doesn't expose it.
- **Graceful errors**: All exceptions are caught and returned as `{"error": "...", "traceback": "..."}` JSON dicts — never crash the server.
- **1-indexed**: Paragraph and table indices match what the user sees in Word.

## License

MIT
