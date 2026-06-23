"""
Word Live MCP Server
====================
Attaches to a *running* Microsoft Word instance via COM (pywin32) and exposes
read/edit/screenshot tools over stdio MCP so Hermes Agent can edit the
document live on screen — mirroring the PowerPoint Live MCP architecture.

Requirements
------------
- Windows + Microsoft Word installed.
- pywin32 + mcp installed in the Python used to run the server.

Register in ~/.hermes/config.yaml:
    mcp_servers:
      word:
        command: "C:/Users/You/AppData/Local/Programs/Python/Python313/python.exe"
        args: ["C:/path/to/word-live-mcp/server.py"]
        connect_timeout: 30
        enabled: true
        timeout: 90

Notes
-----
- Word COM Paragraphs and Tables are 1-indexed to match what the user sees.
- COM objects are apartment-threaded. Each tool call re-initialises COM and
  re-dispatches Word.Application, which transparently reconnects to the
  already-running instance.
- All tools return JSON-serialisable strings.
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import traceback
from contextlib import contextmanager
from functools import wraps
from typing import Any

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP


# ---------------------------------------------------------------------------
# COM session management (identical pattern to ppt-live-mcp)
# ---------------------------------------------------------------------------

@contextmanager
def com_session():
    """Initialise COM for the current thread and yield a fresh app handle."""
    pythoncom.CoInitialize()
    try:
        app = win32com.client.Dispatch("Word.Application")
        try:
            yield app
        finally:
            del app
    finally:
        pythoncom.CoUninitialize()


def word_tool(fn):
    """Decorator: run a tool body inside a COM session with robust error capture.

    Strips the leading ``app`` parameter from the function signature so FastMCP
    does not expose it as a tool argument — it is injected by the decorator at
    runtime from the COM session.
    """
    import inspect

    real_sig = inspect.signature(fn)
    visible_params = [
        p for name, p in real_sig.parameters.items() if name != "app"
    ]
    masked_sig = real_sig.replace(parameters=visible_params)

    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            with com_session() as app:
                return fn(app, *args, **kwargs)
        except Exception as e:
            tb = traceback.format_exc()
            return json.dumps({"error": f"{type(e).__name__}: {e}", "traceback": tb})

    wrapper.__signature__ = masked_sig  # FastMCP reads this
    return wrapper


# ---------------------------------------------------------------------------
# Document helpers
# ---------------------------------------------------------------------------

def _resolve_doc(app, name: str | None):
    """Return the requested Document. If name is None, return ActiveDocument."""
    if name:
        for i in range(1, app.Documents.Count + 1):
            doc = app.Documents.Item(i)
            if doc.Name == name or doc.Name == name + ".docx" or (doc.Path and doc.FullName == name):
                return doc
        raise ValueError(f"No open document matching '{name}'.")
    doc = app.ActiveDocument
    if doc is None:
        raise RuntimeError("No active document. Open a document in Word first.")
    return doc


def _paragraph_summary(para, index: int) -> dict[str, Any]:
    """Extract a compact, JSON-safe description of a paragraph."""
    info: dict[str, Any] = {"index": index}
    try:
        info["text"] = para.Range.Text.replace("\r", "\n").rstrip("\n")
    except Exception:
        info["text"] = ""
    try:
        info["style"] = str(para.Style.NameLocal)
    except Exception:
        info["style"] = None
    try:
        info["alignment"] = para.Alignment
    except Exception:
        pass
    try:
        info["outline_level"] = para.OutlineLevel
    except Exception:
        pass
    return info


def _table_summary(table, index: int) -> dict[str, Any]:
    """Extract a compact summary of a table."""
    info: dict[str, Any] = {
        "index": index,
        "rows": table.Rows.Count,
        "columns": table.Columns.Count,
    }
    try:
        cells = []
        for r in range(1, min(table.Rows.Count + 1, 50)):
            row_cells = []
            for c in range(1, table.Columns.Count + 1):
                try:
                    cell_text = table.Cell(r, c).Range.Text.replace("\r", "").replace("\x07", "").strip()
                    row_cells.append(cell_text)
                except Exception:
                    row_cells.append("(merged)")
            cells.append(row_cells)
        info["cells"] = cells
    except Exception:
        info["cells"] = []
    return info


# ---------------------------------------------------------------------------
# WdEnumerations — constants we need
# ---------------------------------------------------------------------------

WD_ALIGN_LEFT = 0
WD_ALIGN_CENTER = 1
WD_ALIGN_RIGHT = 2
WD_ALIGN_JUSTIFY = 3

WD_COLOR_YELLOW = 10092543   # wdColorYellow
WD_COLOR_RED = 255           # wdColorRed

# Highlight Color Index values:
# wdAutoColor=0, wdBlack=1, wdBlue=2, wdTurquoise=3, wdBrightGreen=4,
# wdPink=5, wdRed=6, wdYellow=7, wdWhite=8, wdDarkBlue=9, wdTeal=10,
# wdGreen=11, wdViolet=12, wdDarkRed=13, wdDarkYellow=14, wdBrown=15,
# wdOliveGreen=16, wdDarkGreen=17, wdDarkCyan=18
WD_COLOR_INDEX_YELLOW = 7    # wdYellow
WD_COLOR_INDEX_RED = 6       # wdRed

WD_FIND_STOP = 0
WD_FIND_CONTINUE = 1

WD_OUTLINE_BODY = 10


# ---------------------------------------------------------------------------
# MCP server
# ---------------------------------------------------------------------------

mcp = FastMCP("word-live")


# === Session / document discovery ==========================================

@mcp.tool()
@word_tool
def list_documents(app) -> str:
    """List all documents currently open in Word.

    Returns name, full path, word count, and whether each is the active document.
    """
    docs = []
    active = None
    try:
        active = app.ActiveDocument
    except Exception:
        active = None
    active_name = active.Name if active else None
    for i in range(1, app.Documents.Count + 1):
        doc = app.Documents.Item(i)
        docs.append({
            "name": doc.Name,
            "path": doc.FullName if doc.Path else "(unsaved)",
            "word_count": doc.Words.Count,
            "active": doc.Name == active_name,
        })
    return json.dumps({"documents": docs, "active": active_name, "count": len(docs)})


@mcp.tool()
@word_tool
def get_active_document(app) -> str:
    """Return details of the currently active document."""
    doc = app.ActiveDocument
    if doc is None:
        return json.dumps({"error": "No active document."})
    current_page = None
    try:
        if app.Selection:
            current_page = app.Selection.Information(3)  # wdActiveEndPageNumber = 3
    except Exception:
        pass
    return json.dumps({
        "name": doc.Name,
        "path": doc.FullName if doc.Path else "(unsaved)",
        "word_count": doc.Words.Count,
        "paragraph_count": doc.Paragraphs.Count,
        "table_count": doc.Tables.Count,
        "section_count": doc.Sections.Count,
        "current_page": current_page,
    })


# === Read ==================================================================

@mcp.tool()
@word_tool
def get_document_content(
    app,
    doc: str | None = None,
    start_paragraph: int = 1,
    max_paragraphs: int = 100,
    include_tables: bool = True,
) -> str:
    """Return paragraphs from the document with text, style, and outline level.

    start_paragraph is 1-indexed. max_paragraphs caps the response.
    include_tables adds a separate tables summary.
    """
    document = _resolve_doc(app, doc)
    total_paras = document.Paragraphs.Count
    end = min(start_paragraph + max_paragraphs - 1, total_paras)
    paragraphs = []
    for i in range(start_paragraph, end + 1):
        para = document.Paragraphs(i)
        summary = _paragraph_summary(para, i)
        # Only include non-empty paragraphs to keep response lean
        if summary.get("text") or summary.get("style", "").startswith("Heading"):
            paragraphs.append(summary)
    result: dict[str, Any] = {
        "document": document.Name,
        "total_paragraphs": total_paras,
        "returned_paragraphs": len(paragraphs),
        "range": f"{start_paragraph}-{end}",
        "paragraphs": paragraphs,
    }
    if include_tables and document.Tables.Count > 0:
        tables = []
        for i in range(1, document.Tables.Count + 1):
            tables.append(_table_summary(document.Tables(i), i))
        result["tables"] = tables
    return json.dumps(result, indent=2)


@mcp.tool()
@word_tool
def get_paragraph(app, paragraph_index: int, doc: str | None = None) -> str:
    """Return the full text and formatting of a specific paragraph (1-indexed)."""
    document = _resolve_doc(app, doc)
    if not (1 <= paragraph_index <= document.Paragraphs.Count):
        return json.dumps({"error": f"paragraph_index {paragraph_index} out of range (1..{document.Paragraphs.Count})"})
    para = document.Paragraphs(paragraph_index)
    info = _paragraph_summary(para, paragraph_index)
    # Font info from first run
    try:
        rng = para.Range
        info["font_name"] = rng.Font.Name
        info["font_size"] = rng.Font.Size
        info["font_bold"] = rng.Font.Bold
        info["font_italic"] = rng.Font.Italic
        info["font_color"] = rng.Font.Color
        info["highlight_color"] = rng.HighlightColorIndex
    except Exception:
        pass
    return json.dumps(info, indent=2)


@mcp.tool()
@word_tool
def get_table(app, table_index: int, doc: str | None = None, max_rows: int = 100) -> str:
    """Return all cell content from a table (1-indexed)."""
    document = _resolve_doc(app, doc)
    if not (1 <= table_index <= document.Tables.Count):
        return json.dumps({"error": f"table_index {table_index} out of range (1..{document.Tables.Count})"})
    return json.dumps(_table_summary(document.Tables(table_index), table_index), indent=2)


@mcp.tool()
@word_tool
def get_selection(app) -> str:
    """Describe the current selection (text, range, or cursor position)."""
    sel = app.Selection
    if not sel:
        return json.dumps({"error": "No active selection."})
    out: dict[str, Any] = {
        "start": sel.Start,
        "end": sel.End,
        "length": sel.End - sel.Start,
    }
    try:
        out["text"] = sel.Text.replace("\r", "\n") if sel.Text else ""
    except Exception:
        out["text"] = ""
    try:
        out["page"] = sel.Information(3)  # wdActiveEndPageNumber
    except Exception:
        pass
    try:
        out["font_name"] = sel.Font.Name
        out["font_size"] = sel.Font.Size
        out["font_bold"] = sel.Font.Bold
        out["font_color"] = sel.Font.Color
        out["highlight_color"] = sel.HighlightColorIndex
    except Exception:
        pass
    return json.dumps(out, indent=2)


@mcp.tool()
@word_tool
def find_text(
    app,
    find: str,
    doc: str | None = None,
    match_case: bool = False,
    whole_word: bool = False,
    max_results: int = 50,
) -> str:
    """Search the document for text and return paragraph indices and positions."""
    document = _resolve_doc(app, doc)
    results = []
    rng = document.Content
    find_obj = rng.Find
    find_obj.ClearFormatting()
    find_obj.Text = find
    find_obj.MatchCase = match_case
    find_obj.MatchWholeWord = whole_word
    find_obj.Forward = True
    find_obj.Wrap = WD_FIND_STOP  # don't wrap around
    while find_obj.Execute() and len(results) < max_results:
        start = rng.Start
        end = rng.End
        # Determine which paragraph this falls in
        para_num = document.Range(0, start).Paragraphs.Count
        results.append({
            "paragraph": para_num,
            "start": start,
            "end": end,
            "found_text": rng.Text[:200],
        })
        # Move range forward to search after this match
        rng.Start = end
        rng.End = document.Content.End
        find_obj = rng.Find
        find_obj.ClearFormatting()
        find_obj.Text = find
        find_obj.MatchCase = match_case
        find_obj.MatchWholeWord = whole_word
        find_obj.Forward = True
        find_obj.Wrap = WD_FIND_STOP
    return json.dumps({
        "document": document.Name,
        "search": find,
        "results_count": len(results),
        "results": results,
    }, indent=2)


# === Edit ==================================================================

@mcp.tool()
@word_tool
def set_paragraph_text(
    app,
    paragraph_index: int,
    text: str,
    doc: str | None = None,
    preserve_format: bool = True,
) -> str:
    """Set the text of a paragraph (1-indexed). Replaces the entire paragraph content.

    preserve_format=True keeps the original paragraph's font/style (default).
    """
    document = _resolve_doc(app, doc)
    if not (1 <= paragraph_index <= document.Paragraphs.Count):
        return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
    para = document.Paragraphs(paragraph_index)
    rng = para.Range
    # Replace text (exclude trailing paragraph mark)
    rng.End = rng.End - 1
    rng.Text = text
    return json.dumps({
        "ok": True,
        "document": document.Name,
        "paragraph": paragraph_index,
        "chars": len(text),
    })


@mcp.tool()
@word_tool
def insert_paragraph_after(
    app,
    paragraph_index: int,
    text: str,
    doc: str | None = None,
    style: str | None = None,
) -> str:
    """Insert a new paragraph AFTER the given paragraph index. Returns the new index."""
    document = _resolve_doc(app, doc)
    if not (1 <= paragraph_index <= document.Paragraphs.Count):
        return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
    para = document.Paragraphs(paragraph_index)
    rng = para.Range
    # Insert at end of this paragraph (after the paragraph mark)
    insert_point = rng.End
    new_rng = document.Range(insert_point, insert_point)
    new_rng.InsertBefore(text + "\r")
    # The new paragraph is at paragraph_index + 1
    new_index = paragraph_index + 1
    if style:
        try:
            document.Paragraphs(new_index).Style = document.Styles(style)
        except Exception:
            pass  # style may not exist
    return json.dumps({
        "ok": True,
        "document": document.Name,
        "inserted_after": paragraph_index,
        "new_paragraph_index": new_index,
    })


@mcp.tool()
@word_tool
def delete_paragraph(app, paragraph_index: int, doc: str | None = None) -> str:
    """Delete a paragraph by index (1-indexed). Use with caution."""
    document = _resolve_doc(app, doc)
    if not (1 <= paragraph_index <= document.Paragraphs.Count):
        return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
    para = document.Paragraphs(paragraph_index)
    rng = para.Range
    # Include the paragraph mark in the deletion
    rng.Delete()
    return json.dumps({
        "ok": True,
        "document": document.Name,
        "deleted": paragraph_index,
        "remaining_paragraphs": document.Paragraphs.Count,
    })


@mcp.tool()
@word_tool
def replace_text(
    app,
    find: str,
    replace: str,
    doc: str | None = None,
    match_case: bool = False,
    whole_word: bool = False,
) -> str:
    """Find-and-replace text throughout the entire document.

    Returns the number of replacements made.
    """
    document = _resolve_doc(app, doc)
    rng = document.Content
    find_obj = rng.Find
    find_obj.ClearFormatting()
    find_obj.Replacement.ClearFormatting()
    find_obj.Text = find
    find_obj.Replacement.Text = replace
    find_obj.MatchCase = match_case
    find_obj.MatchWholeWord = whole_word
    find_obj.Forward = True
    find_obj.Wrap = 1  # wdFindContinue
    find_obj.Format = False
    # wdReplaceAll = 2
    find_obj.Replacement.Wrap = 1
    count = find_obj.Execute(Replace=2)
    return json.dumps({
        "ok": True,
        "document": document.Name,
        "find": find,
        "replace": replace,
        "replacements": count,
    })


# === Formatting ============================================================

@mcp.tool()
@word_tool
def set_font(
    app,
    paragraph_index: int | None = None,
    doc: str | None = None,
    size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color_rgb: str | None = None,
    font_name: str | None = None,
    apply_to: str = "paragraph",  # "paragraph" | "selection"
) -> str:
    """Set font properties on a paragraph or the current selection.

    color_rgb is a hex string like 'FF0000' (red) or '0000FF' (blue).
    apply_to: 'paragraph' uses paragraph_index; 'selection' applies to current selection.
    """
    document = _resolve_doc(app, doc)
    if apply_to == "selection":
        rng = app.Selection.Range if app.Selection else None
        if rng is None:
            return json.dumps({"error": "No active selection."})
    else:
        if paragraph_index is None:
            return json.dumps({"error": "paragraph_index required when apply_to='paragraph'"})
        if not (1 <= paragraph_index <= document.Paragraphs.Count):
            return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
        rng = document.Paragraphs(paragraph_index).Range
    font = rng.Font
    if size is not None:
        font.Size = size
    if bold is not None:
        font.Bold = -1 if bold else 0
    if italic is not None:
        font.Italic = -1 if italic else 0
    if font_name is not None:
        font.Name = font_name
    if color_rgb is not None:
        hex_clean = color_rgb.lstrip("#")
        r = int(hex_clean[0:2], 16)
        g = int(hex_clean[2:4], 16)
        b = int(hex_clean[4:6], 16)
        font.Color = r + g * 256 + b * 65536
    return json.dumps({"ok": True, "applied_to": apply_to, "paragraph": paragraph_index})


@mcp.tool()
@word_tool
def set_highlight(
    app,
    paragraph_index: int | None = None,
    doc: str | None = None,
    color: str = "yellow",  # yellow | red | green | turquoise | none
    apply_to: str = "paragraph",  # "paragraph" | "selection"
) -> str:
    """Set highlight color on a paragraph or selection.

    color options: yellow, red, green, turquoise, pink, none (to remove highlight).
    """
    document = _resolve_doc(app, doc)
    if apply_to == "selection":
        rng = app.Selection.Range if app.Selection else None
        if rng is None:
            return json.dumps({"error": "No active selection."})
    else:
        if paragraph_index is None:
            return json.dumps({"error": "paragraph_index required when apply_to='paragraph'"})
        if not (1 <= paragraph_index <= document.Paragraphs.Count):
            return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
        rng = document.Paragraphs(paragraph_index).Range
    color_map = {
        "yellow": WD_COLOR_INDEX_YELLOW,
        "red": WD_COLOR_INDEX_RED,
        "green": 11,
        "turquoise": 3,
        "pink": 5,
        "blue": 2,
        "none": 0,       # wdAutoColor
        "auto": 0,
    }
    rng.HighlightColorIndex = color_map.get(color.lower(), WD_COLOR_INDEX_YELLOW)
    return json.dumps({"ok": True, "color": color, "applied_to": apply_to, "paragraph": paragraph_index})


@mcp.tool()
@word_tool
def set_paragraph_style(
    app,
    paragraph_index: int,
    style: str,
    doc: str | None = None,
) -> str:
    """Apply a named style to a paragraph. E.g. 'Heading 1', 'Normal', 'List Bullet'."""
    document = _resolve_doc(app, doc)
    if not (1 <= paragraph_index <= document.Paragraphs.Count):
        return json.dumps({"error": f"paragraph_index {paragraph_index} out of range"})
    try:
        document.Paragraphs(paragraph_index).Style = document.Styles(style)
    except Exception as e:
        return json.dumps({"error": f"Style '{style}' not found: {e}"})
    return json.dumps({"ok": True, "paragraph": paragraph_index, "style": style})


# === Tables ================================================================

@mcp.tool()
@word_tool
def set_table_cell(
    app,
    table_index: int,
    row: int,
    column: int,
    text: str,
    doc: str | None = None,
) -> str:
    """Set the text of a specific cell in a table (all 1-indexed)."""
    document = _resolve_doc(app, doc)
    if not (1 <= table_index <= document.Tables.Count):
        return json.dumps({"error": f"table_index {table_index} out of range"})
    table = document.Tables(table_index)
    try:
        cell = table.Cell(row, column)
        rng = cell.Range
        # Cell range includes end-of-cell marker; trim it
        rng.End = rng.End - 1
        rng.Text = text
    except Exception as e:
        return json.dumps({"error": f"Cannot access cell ({row}, {column}): {e}"})
    return json.dumps({
        "ok": True,
        "table": table_index,
        "row": row,
        "column": column,
        "text": text[:100],
    })


@mcp.tool()
@word_tool
def insert_table(
    app,
    rows: int,
    columns: int,
    doc: str | None = None,
    paragraph_index: int | None = None,
) -> str:
    """Insert a table into the document. If paragraph_index is given, inserts after it;
    otherwise inserts at the end of the document."""
    document = _resolve_doc(app, doc)
    if paragraph_index and 1 <= paragraph_index <= document.Paragraphs.Count:
        rng = document.Paragraphs(paragraph_index).Range
        insert_point = rng.End
    else:
        rng = document.Content
        insert_point = rng.End - 1  # before final paragraph mark
    insert_rng = document.Range(insert_point, insert_point)
    insert_rng.InsertAfter("\r")
    insert_point += 1
    table_rng = document.Range(insert_point, insert_point)
    table = document.Tables.Add(table_rng, rows, columns)
    return json.dumps({
        "ok": True,
        "table_index": document.Tables.Count,
        "rows": rows,
        "columns": columns,
    })


# === Visual / export =======================================================

@mcp.tool()
@word_tool
def export_pdf(
    app,
    output_path: str | None = None,
    doc: str | None = None,
) -> str:
    """Export the document to PDF. Returns the file path.
    If output_path is None, saves to a temp file."""
    document = _resolve_doc(app, doc)
    if output_path is None:
        fd, output_path = tempfile.mkstemp(prefix="word_export_", suffix=".pdf")
        os.close(fd)
    # wdExportFormatPDF = 17
    document.ExportAsFixedFormat(
        OutputFileName=output_path,
        ExportFormat=17,
        OpenAfterExport=False,
        OptimizeFor=0,  # wdExportOptimizeForPrint
        Range=0,        # wdExportAllDocument
        Item=0,         # wdExportDocumentContent
        IncludeDocProps=True,
        KeepIRM=True,
        CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
        DocStructureTags=True,
        BitmapMissingFonts=True,
        UseISO19005_1=False,
    )
    size_kb = round(os.path.getsize(output_path) / 1024, 1)
    return json.dumps({"ok": True, "path": output_path, "size_kb": size_kb})


@mcp.tool()
@word_tool
def export_page_image(
    app,
    page: int | None = None,
    output_path: str | None = None,
    doc: str | None = None,
    zoom: int = 200,
) -> str:
    """Export a single page (or current page) as a PNG image for visual QA.
    Uses PDF export + render. page is 1-indexed; None = current page."""
    import fitz  # pymupdf for PDF->PNG
    document = _resolve_doc(app, doc)
    if page is None:
        if app.Selection:
            page = app.Selection.Information(3)  # wdActiveEndPageNumber
        else:
            page = 1
    # Export to temp PDF first
    fd, pdf_path = tempfile.mkstemp(prefix="word_page_", suffix=".pdf")
    os.close(fd)
    document.ExportAsFixedFormat(
        OutputFileName=pdf_path, ExportFormat=17, OpenAfterExport=False,
        OptimizeFor=0, Range=0, Item=0,
        IncludeDocProps=True, KeepIRM=True, CreateBookmarks=0,
        DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False,
    )
    # Render the specific page
    pdf_doc = fitz.open(pdf_path)
    if page < 1 or page > len(pdf_doc):
        pdf_doc.close()
        os.unlink(pdf_path)
        return json.dumps({"error": f"page {page} out of range (1..{len(pdf_doc)})"})
    pdf_page = pdf_doc[page - 1]
    matrix = fitz.Matrix(zoom / 100, zoom / 100)
    pix = pdf_page.get_pixmap(matrix=matrix)
    if output_path is None:
        fd2, output_path = tempfile.mkstemp(prefix=f"word_p{page:02d}_", suffix=".png")
        os.close(fd2)
    pix.save(output_path)
    pdf_doc.close()
    os.unlink(pdf_path)
    size_kb = round(os.path.getsize(output_path) / 1024, 1)
    return json.dumps({
        "ok": True,
        "page": page,
        "path": output_path,
        "size_kb": size_kb,
        "dimensions": f"{pix.width}x{pix.height}",
    })


@mcp.tool()
@word_tool
def update_toc(app, doc: str | None = None) -> str:
    """Update all Tables of Contents in the document."""
    document = _resolve_doc(app, doc)
    updated = 0
    if document.TablesOfContents.Count > 0:
        for i in range(1, document.TablesOfContents.Count + 1):
            document.TablesOfContents(i).Update()
            updated += 1
    return json.dumps({"ok": True, "updated_tocs": updated, "document": document.Name})


@mcp.tool()
@word_tool
def save_document(app, doc: str | None = None) -> str:
    """Save the document. Does NOT save as a new file."""
    document = _resolve_doc(app, doc)
    document.Save()
    return json.dumps({"ok": True, "document": document.Name, "path": document.FullName})


# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
