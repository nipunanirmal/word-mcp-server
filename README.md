# Word Document MCP Server

A powerful, file-based MCP (Model Context Protocol) server for creating and manipulating `.docx` files — no live Word instance required. All operations work directly on `.docx` files using `python-docx`.

## Attribution

This server is a consolidated integration of three excellent open-source Word MCP servers:

- **Base server**: [RenatoTadeuFigueiredo/word-mcp-server](https://github.com/RenatoTadeuFigueiredo/word-mcp-server) — comprehensive file-based document lifecycle, editing, formatting, tables, images, headers/footers, styles, TOC, hyperlinks, comments, bookmarks, and batch builder.
- **Extended tools**: [GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server) — footnotes/endnotes, advanced table formatting (alternating rows, cell shading, alignment, padding, column widths), comment extraction, document protection (password encryption), PDF export, document merge, standalone list tools.
- **Tracked changes, anchors & advanced formatting**: [knorq-ai/docx-mcp-server](https://github.com/knorq-ai/docx-mcp-server) — tracked changes (accept all/reject all), text highlighting, stable paragraph anchors (w14:paraId), list images, threaded comment replies, delete comments, get page layout, search-and-format text, bulk headings, paragraph format introspection, fallback error handling.

All credit for the integrated features goes to the respective authors. This project merely combines their work into a single server.

## How It Works

The server operates entirely on `.docx` files on disk using `python-docx` and `lxml`. No Microsoft Word installation is required. Open a document with `open_document`, make changes with the various tools, and changes are saved back to the file.

## Tools (92)

### Document Lifecycle
| Tool | Description |
|------|-------------|
| `open_document` | Open an existing `.docx` file |
| `create_new_document` | Create a new document |
| `save_document` | Save the current document |
| `close_document` | Close the current document |
| `duplicate_document` | Duplicate the current document |

### Reading & Inspection
| Tool | Description |
|------|-------------|
| `get_document_info` | Document metadata and structure summary |
| `read_document` | Read full document content |
| `get_outline` | Document heading outline |
| `find_text` | Search for text in the document |
| `read_table` | Read table contents |
| `list_styles` | List available styles |
| `get_document_xml` | Get raw XML of the document |
| `read_headers_footers` | Read header/footer content |

### Text Editing
| Tool | Description |
|------|-------------|
| `replace_text` | Find and replace text |
| `replace_text_batch` | Batch find-and-replace: multiple replacements in one call |
| `replace_paragraph` | Replace paragraph content |
| `replace_section` | Replace content under a heading |
| `insert_paragraph` | Insert a new paragraph |
| `delete_paragraph` | Delete a paragraph |
| `delete_section` | Delete a section under a heading |
| `move_section` | Move a section to a new position |

### Formatting
| Tool | Description |
|------|-------------|
| `format_paragraph` | Apply paragraph formatting |
| `format_text_run` | Format a specific text run |
| `add_heading` | Add a heading paragraph |
| `create_style` | Create a custom paragraph style |
| `highlight_text` | Highlight all occurrences of specific text |

### Tables
| Tool | Description |
|------|-------------|
| `insert_table` | Insert a new table |
| `edit_table_cell` | Edit a cell's content |
| `add_table_row` | Add a row to a table |
| `merge_table_cells` | Merge cells in a table |
| `format_table_cell` | Format a cell (bold, color, etc.) |
| `delete_table_row` | Delete a table row |
| `delete_table` | Delete an entire table |
| `format_table` | Format table with borders, header, shading |
| `set_table_cell_shading` | Apply background color to a cell |
| `apply_table_alternating_rows` | Apply alternating row colors |
| `highlight_table_header` | Highlight header row with colors |
| `set_table_cell_alignment` | Set cell text alignment (horizontal + vertical) |
| `set_table_column_width` | Set a single column width |
| `set_table_column_widths` | Set multiple column widths at once |
| `set_table_cell_padding` | Set cell padding/margins |

### Headers & Footers
| Tool | Description |
|------|-------------|
| `set_header` | Set header text |
| `set_footer` | Set footer text |
| `add_image_to_header` | Add image to header |
| `clear_header` | Clear header content |
| `clear_footer` | Clear footer content |

### Images & Layout
| Tool | Description |
|------|-------------|
| `insert_image` | Insert an image |
| `set_page_margins` | Set page margins |
| `set_page_orientation` | Set page orientation |
| `set_page_size` | Set page size |
| `add_section_break` | Add a section break |
| `add_page_break` | Add a page break |
| `set_columns` | Set multi-column layout |

### Lists
| Tool | Description |
|------|-------------|
| `insert_bulleted_list` | Insert a bulleted list |
| `insert_numbered_list` | Insert a numbered list |

### Comments
| Tool | Description |
|------|-------------|
| `add_comment` | Add a native Word comment |
| `get_all_comments` | Extract all comments with metadata |
| `get_comments_by_author` | Filter comments by author |
| `reply_to_comment` | Reply to an existing comment (threaded) |
| `delete_comment` | Delete a comment by ID |

### Tracked Changes
| Tool | Description |
|------|-------------|
| `accept_all_changes` | Accept all tracked changes |
| `reject_all_changes` | Reject all tracked changes |
| `has_tracked_changes` | Check if document has tracked changes |

### Footnotes & Endnotes
| Tool | Description |
|------|-------------|
| `add_footnote` | Add a footnote to a paragraph |
| `read_footnotes` | Read all footnotes |
| `add_endnote` | Add an endnote to a paragraph |

### Document Protection
| Tool | Description |
|------|-------------|
| `protect_document` | Encrypt document with password |
| `unprotect_document` | Decrypt password-protected document |

### Export & Merge
| Tool | Description |
|------|-------------|
| `convert_to_pdf` | Convert document to PDF |
| `merge_documents` | Merge multiple `.docx` files |

### Stable Anchors
| Tool | Description |
|------|-------------|
| `ensure_anchors` | Seed/repair stable w14:paraId anchors on all paragraphs |

### Images
| Tool | Description |
|------|-------------|
| `insert_image` | Insert an image |
| `list_images` | List all embedded images with dimensions |

### Page Layout
| Tool | Description |
|------|-------------|
| `get_page_layout` | Get page size, margins, orientation, detected preset |

### Advanced Formatting
| Tool | Description |
|------|-------------|
| `format_text` | Search-and-format: bold/italic/underline/strikethrough/font/size/color |
| `set_headings` | Bulk convert paragraphs to headings |
| `get_paragraph_format` | Introspect paragraph formatting (style, alignment, spacing, runs) |

### Other
| Tool | Description |
|------|-------------|
| `insert_toc` | Insert table of contents |
| `add_hyperlink` | Add a hyperlink |
| `add_bookmark` | Add a bookmark |
| `set_document_properties` | Set document properties |
| `apply_xml_patch` | Apply raw XML patch |
| `build_document` | Build document from JSON spec |

### Bibliography / Sources
| Tool | Description |
|------|-------------|
| `add_bibliography_source` | Add a single citation source to Word's master bibliography (Sources.xml) |
| `add_bibliography_sources_batch` | Add multiple citation sources in one call (bulk import) |
| `list_bibliography_sources` | List all sources in Word's master bibliography |
| `clear_bibliography_sources` | Remove all sources from the master bibliography |
| `insert_citation` | Replace typed citation text with a Word CITATION field linked to source(s) |
| `insert_citations_batch` | Replace multiple typed citations in one call (batch version of insert_citation) |
| `copy_sources_to_current_list` | Copy sources from master bibliography into the document's Current List |
| `insert_bibliography` | Insert an auto-updating BIBLIOGRAPHY field at the end of the document |

## Requirements

- **Python 3.11+**
- **Packages:**
  ```bash
  pip install python-docx lxml mcp msoffcrypto-tool docx2pdf
  ```
- **For PDF export on Linux/macOS**: Install [LibreOffice](https://www.libreoffice.org/)
- **For PDF export on Windows**: `docx2pdf` uses Microsoft Word via COM

## Installation

### Claude Desktop

Add to `claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "word": {
      "command": "python",
      "args": ["C:/path/to/word-mcp/server.py"]
    }
  }
}
```

### Generic MCP Client

The server communicates over **stdio** using the MCP protocol. Any MCP-compatible client can connect by spawning the server process.

## Usage

Open a document first with `open_document`, then use any of the tools. Changes are saved automatically.

Paragraph and table indices are **1-indexed**.

### Example workflows

- **"Open my document"** → `open_document(path="C:/docs/report.docx")`
- **"Replace 'old term' with 'new term'"** → `replace_text(find="old term", replace="new term")`
- **"Highlight all mentions of 'confidential'"** → `highlight_text(search_text="confidential", color="yellow")`
- **"Accept all tracked changes"** → `accept_all_changes()`
- **"Add alternating row colors to table 1"** → `apply_table_alternating_rows(table_index=1)`
- **"Convert to PDF"** → `convert_to_pdf()`
- **"Merge three documents"** → `merge_documents(source_paths=["a.docx", "b.docx", "c.docx"], output_path="merged.docx")`
- **"Seed stable anchors"** → `ensure_anchors()` — assigns w14:paraId to every paragraph
- **"Make all 'Important' text bold red"** → `format_text(search="Important", bold=True, font_color="FF0000")`
- **"Set paragraphs 1, 3, 5 as Heading 2"** → `set_headings(headings=[{paragraph_index:1, level:2}, {paragraph_index:3, level:2}, {paragraph_index:5, level:2}])`
- **"Get page layout info"** → `get_page_layout()`
- **"List all images"** → `list_images()`

## License

MIT
