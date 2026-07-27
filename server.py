#!/usr/bin/env python3
"""
Word Document MCP Server
========================
A fully-featured MCP server for creating and manipulating .docx files
via Claude Code integration.

Base implementation adapted from RenatoTadeuFigueiredo/word-mcp-server.

Additional features integrated from:
- GongRzhe/Office-Word-MCP-Server: footnotes, advanced table formatting,
  comment extraction, document protection, PDF export, document merge,
  standalone list tools.
- knorq-ai/docx-mcp-server: tracked changes (accept/reject), text highlighting,
  stable anchors (w14:paraId), list images, threaded comment replies,
  delete comments, get page layout, search-and-format text, bulk headings,
  paragraph format introspection, fallback error handling.

Capabilities:
- Full document lifecycle (create, open, save, close, duplicate)
- Reading & searching content (paragraphs, headings, tables, outline)
- Text editing (replace, insert, delete paragraphs and sections)
- Rich formatting (paragraph styles, character runs, alignment, spacing)
- Headers & footers (with page numbers, date fields, images)
- Images (insert from file, resize, alignment, list embedded images)
- Tables (create, read, edit cells, add rows, merge cells, column widths)
- Advanced table formatting (alternating rows, cell shading, alignment, padding)
- Lists (bullet, numbered, multi-level, standalone list tools)
- Document properties (title, author, subject, keywords)
- Page layout (margins, orientation, page size, columns, get layout info)
- Styles (list, create, modify paragraph/character styles)
- Table of contents generation
- Hyperlinks
- Comments (native Word comments, extraction, threaded replies, delete)
- Bookmarks
- Page breaks & section breaks
- Auto-backup before save
- Batch document builder (JSON-driven)
- Tracked changes (accept all / reject all / check)
- Text highlighting (search and highlight)
- Search-and-format text (bold, italic, underline, strikethrough, font, size, color)
- Footnotes & endnotes (add, read)
- Document protection (password encrypt/decrypt)
- PDF export (cross-platform)
- Document merge (combine multiple .docx files)
- Stable paragraph anchors (w14:paraId seeding, repair, idempotent)
- Bulk heading conversion
- Paragraph format introspection
- Fallback error handling with recovery suggestions
"""

import asyncio
import copy
import io
import json
import os
import platform
import re
import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import nsmap, qn
from docx.shared import Cm, Inches, Pt, RGBColor, Emu
from lxml import etree

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# ---------------------------------------------------------------------------
# Server instance & global state
# ---------------------------------------------------------------------------

server = Server("word-mcp-server")

_current_doc_path: Optional[str] = None
_doc: Optional[Document] = None

# XML namespaces
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _require_doc() -> Document:
    """Raise an error if no document is currently open."""
    if _doc is None:
        raise RuntimeError("No document open. Use 'open_document' first.")
    return _doc


def _reload_in_word(path: str):
    """If Microsoft Word is running and has the file open, reload it via AppleScript."""
    script = f'''
tell application "Microsoft Word"
    set docPath to POSIX file "{path}" as string
    repeat with d in documents
        if full name of d is docPath then
            close d saving no
            open (POSIX file "{path}")
            exit repeat
        end if
    end repeat
end tell
'''
    try:
        subprocess.run(
            ["osascript", "-e", script],
            timeout=5,
            capture_output=True,
        )
    except Exception:
        pass  # silently ignore if Word is not running or AppleScript fails


def _save(backup: bool = True):
    """Save the current document. Optionally create a .bak backup first."""
    doc = _require_doc()
    if backup and _current_doc_path and os.path.exists(_current_doc_path):
        bak_path = _current_doc_path + ".bak"
        shutil.copy2(_current_doc_path, bak_path)
    doc.save(_current_doc_path)
    _reload_in_word(_current_doc_path)


# ---------------------------------------------------------------------------
# Bibliography / Sources helpers
# ---------------------------------------------------------------------------

_BIB_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
_BIB_NSMAP = {"b": _BIB_NS}


def _get_sources_xml_path() -> str:
    """Return the path to Word's master Sources.xml, creating the directory if needed."""
    if platform.system() == "Windows":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
        path = os.path.join(base, "Microsoft", "Bibliography", "Sources.xml")
    elif platform.system() == "Darwin":
        base = os.path.expanduser("~/Library/Application Support/Microsoft")
        path = os.path.join(base, "Bibliography", "Sources.xml")
    else:
        base = os.path.expanduser("~/.local/share/Microsoft")
        path = os.path.join(base, "Bibliography", "Sources.xml")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return path


def _load_sources_tree():
    """Load the Sources.xml tree. Returns (tree, root). Creates an empty one if missing."""
    from lxml import etree as _etree

    path = _get_sources_xml_path()
    if os.path.exists(path):
        tree = _etree.parse(path)
        root = tree.getroot()
    else:
        root = _etree.Element(
            f"{{{_BIB_NS}}}Sources",
            attrib={"SelectedStyle": ""},
            nsmap={"b": _BIB_NS},
        )
        tree = _etree.ElementTree(root)
    return tree, root


def _xml_escape(s: str) -> str:
    if not s:
        return s
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _build_author_xml(authors: list) -> str:
    """Build the <b:Author> XML block from a list of author dicts."""
    persons = []
    corporate = []
    for a in authors:
        if a.get("corporate"):
            corporate.append(
                f"<b:Corporate>{_xml_escape(a['corporate'])}</b:Corporate>"
            )
        else:
            persons.append(
                f"<b:Person>"
                f"<b:Last>{_xml_escape(a.get('last', ''))}</b:Last>"
                f"<b:First>{_xml_escape(a.get('first', ''))}</b:First>"
                f"</b:Person>"
            )
    name_list = "".join(persons) + "".join(corporate)
    return (
        f"<b:Author><b:Author><b:NameList>{name_list}</b:NameList></b:Author></b:Author>"
    )


def _build_source_xml(src: dict) -> str:
    """Build a single <b:Source> XML element string from a source dict."""
    import uuid as _uuid

    source_type = src.get("source_type", "JournalArticle")
    year = src.get("year", "")
    title = src.get("title", "")

    # Auto-generate tag if not provided
    tag = src.get("tag", "")
    if not tag:
        authors = src.get("authors", [])
        if authors:
            first_author = authors[0]
            if first_author.get("corporate"):
                name_part = first_author["corporate"].split()[0]
            else:
                name_part = first_author.get("last", "Source")
        else:
            name_part = "Source"
        # Clean tag: remove non-alphanumeric, capitalize
        name_part = re.sub(r"[^A-Za-z0-9]", "", name_part)
        tag = f"{name_part}{year}"

    guid = src.get("guid", "")
    if not guid:
        guid = str(_uuid.uuid4()).upper()

    parts = [
        "<b:Source>",
        f"<b:Tag>{_xml_escape(tag)}</b:Tag>",
        f"<b:SourceType>{_xml_escape(source_type)}</b:SourceType>",
        f"<b:Guid>{{{_xml_escape(guid)}}}</b:Guid>",
    ]

    # Authors
    authors = src.get("authors", [])
    if authors:
        parts.append(_build_author_xml(authors))

    # Title
    if title:
        parts.append(f"<b:Title>{_xml_escape(title)}</b:Title>")

    # Year
    if year:
        parts.append(f"<b:Year>{_xml_escape(str(year))}</b:Year>")

    # Journal / periodical name
    journal = src.get("journal_name", "")
    if journal:
        parts.append(f"<b:JournalName>{_xml_escape(journal)}</b:JournalName>")

    # Book title (for BookSection)
    book_title = src.get("book_title", "")
    if book_title:
        parts.append(f"<b:BookTitle>{_xml_escape(book_title)}</b:BookTitle>")

    # Volume / Issue / Pages
    if src.get("volume"):
        parts.append(f"<b:Volume>{_xml_escape(str(src['volume']))}</b:Volume>")
    if src.get("issue"):
        parts.append(f"<b:Issue>{_xml_escape(str(src['issue']))}</b:Issue>")
    if src.get("pages"):
        parts.append(f"<b:Pages>{_xml_escape(str(src['pages']))}</b:Pages>")

    # Publisher / City
    if src.get("publisher"):
        parts.append(f"<b:Publisher>{_xml_escape(src['publisher'])}</b:Publisher>")
    if src.get("city"):
        parts.append(f"<b:City>{_xml_escape(src['city'])}</b:City>")

    # DOI / URL
    if src.get("doi"):
        parts.append(f"<b:DOI>{_xml_escape(src['doi'])}</b:DOI>")
    if src.get("url"):
        parts.append(f"<b:URL>{_xml_escape(src['url'])}</b:URL>")

    # Abstract
    if src.get("abstract"):
        parts.append(f"<b:Abstract>{_xml_escape(src['abstract'])}</b:Abstract>")

    # Edition / ISBN / ISSN
    if src.get("edition"):
        parts.append(f"<b:Edition>{_xml_escape(str(src['edition']))}</b:Edition>")
    if src.get("isbn"):
        parts.append(f"<b:ISBN>{_xml_escape(src['isbn'])}</b:ISBN>")
    if src.get("issn"):
        parts.append(f"<b:ISSN>{_xml_escape(src['issn'])}</b:ISSN>")

    # Language / Comments
    if src.get("language"):
        parts.append(f"<b:Language>{_xml_escape(src['language'])}</b:Language>")
    if src.get("comments"):
        parts.append(f"<b:Comments>{_xml_escape(src['comments'])}</b:Comments>")

    # Conference name
    if src.get("conference_name"):
        parts.append(
            f"<b:ConferenceName>{_xml_escape(src['conference_name'])}</b:ConferenceName>"
        )

    # Medium / Number / Country
    if src.get("medium"):
        parts.append(f"<b:Medium>{_xml_escape(src['medium'])}</b:Medium>")
    if src.get("number"):
        parts.append(f"<b:Number>{_xml_escape(str(src['number']))}</b:Number>")
    if src.get("country"):
        parts.append(f"<b:Country>{_xml_escape(src['country'])}</b:Country>")

    # Thesis-specific
    if src.get("institution"):
        parts.append(f"<b:Institution>{_xml_escape(src['institution'])}</b:Institution>")
    if src.get("department"):
        parts.append(f"<b:Department>{_xml_escape(src['department'])}</b:Department>")
    if src.get("thesis_type"):
        parts.append(f"<b:ThesisType>{_xml_escape(src['thesis_type'])}</b:ThesisType>")

    # Access date
    if src.get("access_date"):
        parts.append(f"<b:AccessDate>{_xml_escape(src['access_date'])}</b:AccessDate>")

    # Production company (Film)
    if src.get("production_company"):
        parts.append(
            f"<b:ProductionCompany>{_xml_escape(src['production_company'])}</b:ProductionCompany>"
        )

    parts.append("</b:Source>")
    return "".join(parts)


def _add_sources_to_xml(sources: list) -> tuple:
    """Add a list of source dicts to Sources.xml. Returns (added_count, skipped_count, messages)."""
    from lxml import etree as _etree

    tree, root = _load_sources_tree()

    # Collect existing tags to avoid duplicates
    existing_tags = set()
    for src_elem in root.findall(f"{{{_BIB_NS}}}Source"):
        tag_elem = src_elem.find(f"{{{_BIB_NS}}}Tag")
        if tag_elem is not None and tag_elem.text:
            existing_tags.add(tag_elem.text)

    added = 0
    skipped = 0
    messages = []

    for src in sources:
        tag = src.get("tag", "")
        if not tag:
            authors = src.get("authors", [])
            if authors:
                first_author = authors[0]
                if first_author.get("corporate"):
                    name_part = first_author["corporate"].split()[0]
                else:
                    name_part = first_author.get("last", "Source")
            else:
                name_part = "Source"
            name_part = re.sub(r"[^A-Za-z0-9]", "", name_part)
            tag = f"{name_part}{src.get('year', '')}"
            src["tag"] = tag

        if tag in existing_tags:
            skipped += 1
            messages.append(f"Skipped duplicate tag: {tag}")
            continue

        source_xml = _build_source_xml(src)
        # Wrap with namespace declarations so lxml can parse the b: prefix
        wrapped = (
            f'<b:Sources xmlns:b="{_BIB_NS}">'
            f'{source_xml}'
            f'</b:Sources>'
        )
        wrapped_elem = _etree.fromstring(wrapped)
        new_elem = wrapped_elem[0]  # The <b:Source> child
        root.append(new_elem)
        existing_tags.add(tag)
        added += 1
        messages.append(f"Added: {tag} - {src.get('title', '')[:60]}")

    # Write back
    path = _get_sources_xml_path()
    tree.write(path, xml_declaration=True, encoding="utf-8", standalone=True)

    return added, skipped, messages


def _list_sources_from_xml() -> list:
    """Read all sources from Sources.xml and return a list of dicts with summary info."""
    tree, root = _load_sources_tree()
    sources = []
    for src_elem in root.findall(f"{{{_BIB_NS}}}Source"):
        info = {}
        tag_elem = src_elem.find(f"{{{_BIB_NS}}}Tag")
        info["tag"] = tag_elem.text if tag_elem is not None else ""
        type_elem = src_elem.find(f"{{{_BIB_NS}}}SourceType")
        info["source_type"] = type_elem.text if type_elem is not None else ""
        title_elem = src_elem.find(f"{{{_BIB_NS}}}Title")
        info["title"] = title_elem.text if title_elem is not None else ""
        year_elem = src_elem.find(f"{{{_BIB_NS}}}Year")
        info["year"] = year_elem.text if year_elem is not None else ""
        # Extract first author
        person = src_elem.find(f".//{{{_BIB_NS}}}Person")
        corp = src_elem.find(f".//{{{_BIB_NS}}}Corporate")
        if person is not None:
            last = person.find(f"{{{_BIB_NS}}}Last")
            first = person.find(f"{{{_BIB_NS}}}First")
            last_text = last.text if last is not None else ""
            first_text = first.text if first is not None else ""
            info["first_author"] = f"{last_text}, {first_text}".strip(", ")
        elif corp is not None:
            info["first_author"] = corp.text or ""
        else:
            info["first_author"] = ""
        sources.append(info)
    return sources


def _clear_sources_xml() -> int:
    """Remove all sources from Sources.xml. Returns count removed."""
    tree, root = _load_sources_tree()
    count = len(root.findall(f"{{{_BIB_NS}}}Source"))
    # Remove all Source elements
    for src_elem in root.findall(f"{{{_BIB_NS}}}Source"):
        root.remove(src_elem)
    path = _get_sources_xml_path()
    tree.write(path, xml_declaration=True, encoding="utf-8", standalone=True)
    return count


# ---------------------------------------------------------------------------
# Citation & Bibliography field helpers (in-document)
# ---------------------------------------------------------------------------

def _make_citation_field_runs(tags: list) -> list:
    """Create a list of w:r XML elements representing CITATION fields for the given source tags."""
    runs = []
    for i, tag in enumerate(tags):
        if i > 0:
            r_sep = OxmlElement('w:r')
            t_sep = OxmlElement('w:t')
            t_sep.text = "; "
            t_sep.set(qn('xml:space'), 'preserve')
            r_sep.append(t_sep)
            runs.append(r_sep)

        # Begin field
        r_begin = OxmlElement('w:r')
        fld_begin = OxmlElement('w:fldChar')
        fld_begin.set(qn('w:fldCharType'), 'begin')
        r_begin.append(fld_begin)
        runs.append(r_begin)

        # Instruction
        r_instr = OxmlElement('w:r')
        instr = OxmlElement('w:instrText')
        instr.set(qn('xml:space'), 'preserve')
        instr.text = f' CITATION {tag} \\l 1033 '
        r_instr.append(instr)
        runs.append(r_instr)

        # Separate
        r_sep2 = OxmlElement('w:r')
        fld_sep = OxmlElement('w:fldChar')
        fld_sep.set(qn('w:fldCharType'), 'separate')
        r_sep2.append(fld_sep)
        runs.append(r_sep2)

        # Display text (placeholder — Word updates on F9)
        r_text = OxmlElement('w:r')
        t_text = OxmlElement('w:t')
        t_text.set(qn('xml:space'), 'preserve')
        t_text.text = f"({tag})"
        r_text.append(t_text)
        runs.append(r_text)

        # End field
        r_end = OxmlElement('w:r')
        fld_end = OxmlElement('w:fldChar')
        fld_end.set(qn('w:fldCharType'), 'end')
        r_end.append(fld_end)
        runs.append(r_end)

    return runs


def _make_bibliography_field_runs() -> list:
    """Create a list of w:r XML elements representing a BIBLIOGRAPHY field."""
    runs = []

    r_begin = OxmlElement('w:r')
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    r_begin.append(fld_begin)
    runs.append(r_begin)

    r_instr = OxmlElement('w:r')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' BIBLIOGRAPHY \\l 1033 '
    r_instr.append(instr)
    runs.append(r_instr)

    r_sep = OxmlElement('w:r')
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    r_sep.append(fld_sep)
    runs.append(r_sep)

    r_text = OxmlElement('w:r')
    t_text = OxmlElement('w:t')
    t_text.text = "Bibliography will appear here. Press Ctrl+A then F9 to update."
    r_text.append(t_text)
    runs.append(r_text)

    r_end = OxmlElement('w:r')
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    r_end.append(fld_end)
    runs.append(r_end)

    return runs


def _replace_text_with_citation_in_para(para, search_text: str, tags: list) -> bool:
    """Find search_text in paragraph, replace it with CITATION field runs. Returns True if found."""
    full_text = para.text
    if search_text not in full_text:
        return False

    idx = full_text.find(search_text)
    before = full_text[:idx]
    after = full_text[idx + len(search_text):]

    # Clear all existing runs
    p_elem = para._element
    for r_elem in p_elem.findall(qn('w:r')):
        p_elem.remove(r_elem)

    # Add "before" text
    if before:
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = before
        r.append(t)
        p_elem.append(r)

    # Add citation field runs
    for cf_run in _make_citation_field_runs(tags):
        p_elem.append(cf_run)

    # Add "after" text
    if after:
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = after
        r.append(t)
        p_elem.append(r)

    return True


def _replace_cell_with_citations(cell, tags: list):
    """Replace entire table cell content with CITATION field runs."""
    tc = cell._tc
    for p in tc.findall(qn('w:p')):
        tc.remove(p)
    new_p = OxmlElement('w:p')
    tc.append(new_p)
    for cf_run in _make_citation_field_runs(tags):
        new_p.append(cf_run)


def _copy_sources_to_current_list(doc_path: str) -> int:
    """Copy all sources from master Sources.xml into the document's Current List (customXml/item1.xml).
    Returns the number of sources copied."""
    import zipfile
    from copy import deepcopy

    # 1. Read master Sources.xml
    master_tree, master_root = _load_sources_tree()
    master_sources = master_root.findall(f"{{{_BIB_NS}}}Source")

    # 2. Read the document's customXml/item1.xml
    with zipfile.ZipFile(doc_path, 'r') as z:
        item1_bytes = z.read('customXml/item1.xml')

    doc_root = etree.fromstring(item1_bytes)

    # 3. Remove existing sources
    for src in doc_root.findall(f"{{{_BIB_NS}}}Source"):
        doc_root.remove(src)

    # 4. Add all master sources
    for src in master_sources:
        doc_root.append(deepcopy(src))

    # 5. Serialize
    new_item1 = etree.tostring(doc_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 6. Write back to .docx
    tmp_path = doc_path + ".tmp"
    with zipfile.ZipFile(doc_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'customXml/item1.xml':
                    zout.writestr(item, new_item1)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.replace(tmp_path, doc_path)
    return len(master_sources)


def _para_style(para) -> str:
    return para.style.name if para.style else ""


def _heading_level(para) -> Optional[int]:
    m = re.match(r"Heading (\d+)", _para_style(para))
    return int(m.group(1)) if m else None


def _find_heading_para(doc: Document, heading_text: str):
    """Return the paragraph object matching a heading (case-insensitive partial match)."""
    for p in doc.paragraphs:
        if heading_text.lower() in p.text.lower() and _heading_level(p) is not None:
            return p
    return None


def _find_paras_under_heading(doc: Document, heading_text: str) -> list[int]:
    """
    Return list of paragraph indexes that belong to the section under a heading.
    The section ends when a heading of equal or higher level is encountered.
    """
    paras = doc.paragraphs
    start_idx = None
    heading_lvl = None
    for i, p in enumerate(paras):
        if heading_text.lower() in p.text.lower() and _heading_level(p) is not None:
            start_idx = i
            heading_lvl = _heading_level(p)
            break
    if start_idx is None:
        return []
    result = []
    for i in range(start_idx + 1, len(paras)):
        lvl = _heading_level(paras[i])
        if lvl is not None and lvl <= heading_lvl:
            break
        result.append(i)
    return result


def _delete_paragraph(para):
    """Remove a paragraph element from the document body."""
    p = para._p
    parent = p.getparent()
    if parent is not None:
        parent.remove(p)


def _insert_paragraph_after(ref_para, text: str, style_id: Optional[str] = None):
    """Insert a new w:p element immediately after ref_para."""
    new_p = OxmlElement("w:p")
    if style_id:
        pPr = OxmlElement("w:pPr")
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), style_id)
        pPr.append(pStyle)
        new_p.append(pPr)
    if text:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        new_p.append(r)
    ref_para._p.addnext(new_p)
    return new_p


def _get_style_id(doc: Document, style_name: str) -> Optional[str]:
    """Return the internal style_id for a given style name (case-insensitive)."""
    for s in doc.styles:
        if s.name.lower() == style_name.lower():
            return s.style_id
    return None


def _resolve_style_name(doc: Document, style_name: str) -> Optional[str]:
    """Return the canonical style name if found."""
    for s in doc.styles:
        if s.name.lower() == style_name.lower():
            return s.name
    return None


def _set_run_formatting(run, bold=None, italic=None, underline=None,
                        strike=None, font_size=None, font_name=None,
                        color_hex=None, highlight_color=None,
                        all_caps=None, small_caps=None):
    """Apply character-level formatting to a run."""
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if underline is not None:
        run.underline = underline
    if strike is not None:
        run.font.strike = strike
    if font_size is not None:
        run.font.size = Pt(font_size)
    if font_name is not None:
        run.font.name = font_name
    if color_hex is not None:
        h = color_hex.lstrip("#")
        run.font.color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    if all_caps is not None:
        run.font.all_caps = all_caps
    if small_caps is not None:
        run.font.small_caps = small_caps


def _xml_to_str(element) -> str:
    return etree.tostring(element, pretty_print=True).decode()


def _make_oxml_element(tag: str) -> Any:
    return OxmlElement(tag)


# ---------------------------------------------------------------------------
# Header / Footer helpers
# ---------------------------------------------------------------------------

def _get_or_create_hdrftr(section, which: str, link_to_previous: bool = False):
    """
    Return the header or footer object for a section.
    which: 'header' | 'footer' | 'first_page_header' | 'first_page_footer'
    """
    attr_map = {
        "header": "header",
        "footer": "footer",
        "first_page_header": "first_page_header",
        "first_page_footer": "first_page_footer",
        "even_page_header": "even_page_header",
        "even_page_footer": "even_page_footer",
    }
    hf = getattr(section, attr_map[which])
    hf.is_linked_to_previous = link_to_previous
    return hf


def _add_page_number_field(paragraph, alignment: str = "CENTER",
                            fmt: str = "PAGE_NUMBER"):
    """
    Insert a { PAGE } or { PAGE } / { NUMPAGES } field into a paragraph.
    fmt: 'PAGE_NUMBER' | 'PAGE_OF_PAGES'
    """
    align_map = {
        "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
        "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
        "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    paragraph.alignment = align_map.get(alignment.upper(), WD_ALIGN_PARAGRAPH.CENTER)
    paragraph.clear()

    if fmt == "PAGE_OF_PAGES":
        # Build "Page X of Y"
        run = paragraph.add_run("Page ")
        _insert_fld_char(paragraph, "PAGE")
        paragraph.add_run(" of ")
        _insert_fld_char(paragraph, "NUMPAGES")
    else:
        _insert_fld_char(paragraph, "PAGE")


def _insert_fld_char(paragraph, instr: str):
    """Insert a simple Word field (PAGE, NUMPAGES, DATE, etc.) into a paragraph."""
    run = OxmlElement("w:r")
    # fldChar begin
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run.append(fld_begin)
    paragraph._p.append(run)

    run2 = OxmlElement("w:r")
    instr_el = OxmlElement("w:instrText")
    instr_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_el.text = f" {instr} "
    run2.append(instr_el)
    paragraph._p.append(run2)

    run3 = OxmlElement("w:r")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run3.append(fld_end)
    paragraph._p.append(run3)


# ---------------------------------------------------------------------------
# Table helpers
# ---------------------------------------------------------------------------

def _set_cell_background(cell, hex_color: str):
    """Set a table cell background shading color (hex without #)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#").upper())
    tcPr.append(shd)


def _set_cell_border(cell, top=None, bottom=None, left=None, right=None,
                     color: str = "000000", size: int = 4):
    """Set individual borders on a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    border_map = {"top": top, "bottom": bottom, "left": left, "right": right}
    for side, val in border_map.items():
        if val:
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:val"), val)
            el.set(qn("w:sz"), str(size))
            el.set(qn("w:color"), color.lstrip("#"))
            tcBorders.append(el)
    tcPr.append(tcBorders)


def _merge_cells_horizontal(table, row: int, start_col: int, end_col: int):
    """Merge cells horizontally in a row from start_col to end_col (inclusive)."""
    cell_a = table.cell(row, start_col)
    cell_b = table.cell(row, end_col)
    cell_a.merge(cell_b)


def _merge_cells_vertical(table, col: int, start_row: int, end_row: int):
    """Merge cells vertically in a column from start_row to end_row (inclusive)."""
    cell_a = table.cell(start_row, col)
    cell_b = table.cell(end_row, col)
    cell_a.merge(cell_b)


# ---------------------------------------------------------------------------
# Hyperlink helper
# ---------------------------------------------------------------------------

def _add_hyperlink(paragraph, text: str, url: str,
                   color_hex: str = "0563C1", underline: bool = True):
    """Add a clickable hyperlink run to an existing paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    run_el = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)
    if color_hex:
        color_el = OxmlElement("w:color")
        color_el.set(qn("w:val"), color_hex.lstrip("#"))
        rPr.append(color_el)
    if underline:
        u_el = OxmlElement("w:u")
        u_el.set(qn("w:val"), "single")
        rPr.append(u_el)
    run_el.append(rPr)

    t_el = OxmlElement("w:t")
    t_el.text = text
    t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    run_el.append(t_el)
    hyperlink.append(run_el)
    paragraph._p.append(hyperlink)
    return hyperlink


# ---------------------------------------------------------------------------
# Bookmark helpers
# ---------------------------------------------------------------------------

def _add_bookmark(paragraph, bookmark_name: str, bookmark_id: int = 1):
    """Wrap the paragraph content in a Word bookmark."""
    p = paragraph._p
    bm_start = OxmlElement("w:bookmarkStart")
    bm_start.set(qn("w:id"), str(bookmark_id))
    bm_start.set(qn("w:name"), bookmark_name)
    bm_end = OxmlElement("w:bookmarkEnd")
    bm_end.set(qn("w:id"), str(bookmark_id))
    p.insert(0, bm_start)
    p.append(bm_end)


# ---------------------------------------------------------------------------
# Comment helpers
# ---------------------------------------------------------------------------

_comment_id_counter = 0


def _add_native_comment(doc: Document, paragraph, text: str,
                         author: str = "Claude", initials: str = "AI"):
    """
    Add a native Word comment to a paragraph (appears in the comments panel).
    This injects the required XML into word/comments.xml part.
    """
    global _comment_id_counter
    _comment_id_counter += 1
    cid = _comment_id_counter

    # Access or create the comments part
    try:
        comments_part = doc.part.comments_part
        comments_el = comments_part._element
    except AttributeError:
        # Create comments part from scratch
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        CT_COMMENTS = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '</w:comments>'
        )
        comments_el = etree.fromstring(CT_COMMENTS.encode())
        # Fallback: insert inline text comment
        paragraph.add_run(f" [COMMENT({author}): {text}]")
        return cid

    date_str = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    comment_xml = (
        f'<w:comment xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'w:id="{cid}" w:author="{author}" w:date="{date_str}" w:initials="{initials}">'
        f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'
        f'</w:comment>'
    )
    comment_el = etree.fromstring(comment_xml.encode())
    comments_el.append(comment_el)

    # Wrap paragraph runs in comment range markers
    p = paragraph._p
    cm_start = OxmlElement("w:commentRangeStart")
    cm_start.set(qn("w:id"), str(cid))
    cm_end = OxmlElement("w:commentRangeEnd")
    cm_end.set(qn("w:id"), str(cid))
    cm_ref_run = OxmlElement("w:r")
    cm_ref = OxmlElement("w:commentReference")
    cm_ref.set(qn("w:id"), str(cid))
    cm_ref_run.append(cm_ref)

    p.insert(0, cm_start)
    p.append(cm_end)
    p.append(cm_ref_run)
    return cid


# ---------------------------------------------------------------------------
# Table of Contents helper
# ---------------------------------------------------------------------------

def _insert_toc(doc: Document, title: str = "Table of Contents",
                max_level: int = 3, after_heading: Optional[str] = None):
    """
    Insert a Word TOC field before the first Heading 1, or after a specified
    heading. Word will update the TOC on first open (Ctrl+A, F9).
    """
    def _make_toc_heading_para(doc, title):
        p = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), "TOCHeading")
        pPr.append(pStyle)
        p.append(pPr)
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = title
        r.append(t)
        p.append(r)
        return p

    def _make_toc_field_para(doc, max_level):
        p = OxmlElement("w:p")
        r_begin = OxmlElement("w:r")
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        fld_begin.set(qn("w:dirty"), "true")
        r_begin.append(fld_begin)
        p.append(r_begin)
        r_instr = OxmlElement("w:r")
        instr = OxmlElement("w:instrText")
        instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        instr.text = f' TOC \\o "1-{max_level}" \\h \\z \\u '
        r_instr.append(instr)
        p.append(r_instr)
        r_sep = OxmlElement("w:r")
        fld_sep = OxmlElement("w:fldChar")
        fld_sep.set(qn("w:fldCharType"), "separate")
        r_sep.append(fld_sep)
        p.append(r_sep)
        r_ph = OxmlElement("w:r")
        t_ph = OxmlElement("w:t")
        t_ph.text = "Clique com o botao direito para atualizar o indice."
        r_ph.append(t_ph)
        p.append(r_ph)
        r_end = OxmlElement("w:r")
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end")
        r_end.append(fld_end)
        p.append(r_end)
        return p

    body = doc.element.body

    # at_end sentinel — just append
    if after_heading == "__end__":
        if title:
            doc.add_paragraph(title, style="TOC Heading")
        doc.add_paragraph()._p.extend([_make_toc_field_para(doc, max_level)])
        return

    # Determine the reference element to insert before
    insert_before = None
    if after_heading:
        ref = _find_heading_para(doc, after_heading)
        if ref is not None:
            insert_before = ref._p.getnext()
    if insert_before is None:
        # Default: insert before the first Heading 1
        for child in list(body):
            pStyle = child.find('.//' + qn('w:pStyle'))
            if pStyle is not None and pStyle.get(qn('w:val')) == 'Heading1':
                insert_before = child
                break
    if insert_before is None:
        # Fallback: append at end
        if title:
            doc.add_paragraph(title, style="TOC Heading")
        doc.add_paragraph()._p.extend([_make_toc_field_para(doc, max_level)])
        return

    # Insert before reference: heading first, then field right after heading
    toc_field = _make_toc_field_para(doc, max_level)
    if title:
        toc_head = _make_toc_heading_para(doc, title)
        insert_before.addprevious(toc_field)
        toc_field.addprevious(toc_head)
    else:
        insert_before.addprevious(toc_field)


# ---------------------------------------------------------------------------
# Column layout helper
# ---------------------------------------------------------------------------

def _set_section_columns(section, num_cols: int, spacing_cm: float = 1.25,
                          equal_width: bool = True):
    """Set multi-column layout for a document section."""
    sectPr = section._sectPr
    # Remove existing cols element
    existing = sectPr.find(qn("w:cols"))
    if existing is not None:
        sectPr.remove(existing)
    cols_el = OxmlElement("w:cols")
    cols_el.set(qn("w:num"), str(num_cols))
    cols_el.set(qn("w:space"), str(int(Cm(spacing_cm).emu / 914)))  # EMU -> twips
    cols_el.set(qn("w:equalWidth"), "1" if equal_width else "0")
    sectPr.append(cols_el)


# ---------------------------------------------------------------------------
# Style creation helper
# ---------------------------------------------------------------------------

def _create_paragraph_style(doc: Document, style_name: str,
                              base_style: str = "Normal",
                              font_name: Optional[str] = None,
                              font_size: Optional[float] = None,
                              bold: Optional[bool] = None,
                              italic: Optional[bool] = None,
                              color_hex: Optional[str] = None,
                              alignment: Optional[str] = None,
                              space_before: Optional[float] = None,
                              space_after: Optional[float] = None):
    """Create a new named paragraph style in the document."""
    # Check if already exists
    existing = _resolve_style_name(doc, style_name)
    if existing:
        style = doc.styles[existing]
    else:
        style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

    if base_style:
        try:
            style.base_style = doc.styles[base_style]
        except KeyError:
            pass

    font = style.font
    if font_name:
        font.name = font_name
    if font_size:
        font.size = Pt(font_size)
    if bold is not None:
        font.bold = bold
    if italic is not None:
        font.italic = italic
    if color_hex:
        h = color_hex.lstrip("#")
        font.color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    pf = style.paragraph_format
    align_map = {
        "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
        "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
        "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
        "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    if alignment:
        pf.alignment = align_map.get(alignment.upper(), WD_ALIGN_PARAGRAPH.LEFT)
    if space_before is not None:
        pf.space_before = Pt(space_before)
    if space_after is not None:
        pf.space_after = Pt(space_after)

    return style


# ---------------------------------------------------------------------------
# Tracked changes helpers (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

def _accept_all_changes(doc: Document):
    """Accept all tracked changes in the document by unwrapping w:ins and removing w:del."""
    body = doc.element.body
    # Process all w:ins elements: unwrap (keep content, remove wrapper)
    for ins_el in body.findall('.//' + qn('w:ins')):
        parent = ins_el.getparent()
        if parent is None:
            continue
        for child in list(ins_el):
            ins_el.remove(child)
            parent.insert(list(parent).index(ins_el), child)
        parent.remove(ins_el)
    # Process all w:del elements: remove entirely (content is deleted)
    for del_el in body.findall('.//' + qn('w:del')):
        parent = del_el.getparent()
        if parent is not None:
            parent.remove(del_el)
    # Remove move-tracking markers
    for tag in ['w:moveFrom', 'w:moveTo']:
        for el in body.findall('.//' + qn(tag)):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
    # Clean up paragraph mark revisions (pPr > rPr > w:ins / w:del)
    for pPr in body.findall('.//' + qn('w:pPr')):
        rPr = pPr.find(qn('w:rPr'))
        if rPr is not None:
            for rev_tag in ['w:ins', 'w:del', 'w:rPrChange']:
                for el in rPr.findall(qn(rev_tag)):
                    rPr.remove(el)
    # Clean up run property revisions
    for rPr in body.findall('.//' + qn('w:rPr')):
        for rev_tag in ['w:rPrChange']:
            for el in rPr.findall(qn(rev_tag)):
                rPr.remove(el)


def _reject_all_changes(doc: Document):
    """Reject all tracked changes: remove w:ins content, unwrap w:del (keep content)."""
    body = doc.element.body
    # Process all w:del elements: unwrap (keep deleted text as normal text)
    for del_el in body.findall('.//' + qn('w:del')):
        parent = del_el.getparent()
        if parent is None:
            continue
        for child in list(del_el):
            del_el.remove(child)
            parent.insert(list(parent).index(del_el), child)
        parent.remove(del_el)
    # Process all w:ins elements: remove entirely (content is rejected)
    for ins_el in body.findall('.//' + qn('w:ins')):
        parent = ins_el.getparent()
        if parent is not None:
            parent.remove(ins_el)
    # Remove move-tracking markers
    for tag in ['w:moveFrom', 'w:moveTo']:
        for el in body.findall('.//' + qn(tag)):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
    # Clean up paragraph mark revisions
    for pPr in body.findall('.//' + qn('w:pPr')):
        rPr = pPr.find(qn('w:rPr'))
        if rPr is not None:
            for rev_tag in ['w:ins', 'w:del', 'w:rPrChange']:
                for el in rPr.findall(qn(rev_tag)):
                    rPr.remove(el)
    # Clean up run property revisions
    for rPr in body.findall('.//' + qn('w:rPr')):
        for rev_tag in ['w:rPrChange']:
            for el in rPr.findall(qn(rev_tag)):
                rPr.remove(el)


def _has_tracked_changes(doc: Document) -> bool:
    """Check if the document contains any tracked changes."""
    body = doc.element.body
    for tag in ['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo']:
        if body.findall('.//' + qn(tag)):
            return True
    return False


# ---------------------------------------------------------------------------
# Comment extraction helpers (adapted from GongRzhe/Office-Word-MCP-Server)
# ---------------------------------------------------------------------------

def _extract_all_comments(doc: Document) -> list[dict]:
    """Extract all comments from a Word document's comments part."""
    comments = []
    try:
        document_part = doc.part
        comments_part = None
        for rel_id, rel in document_part.rels.items():
            if 'comments' in rel.reltype and 'comments' == rel.reltype.split('/')[-1]:
                comments_part = rel.target_part
                break
        if comments_part:
            comment_elements = comments_part.element.xpath('.//w:comment')
            for idx, ce in enumerate(comment_elements):
                cid = ce.get(qn('w:id'), str(idx))
                author = ce.get(qn('w:author'), 'Unknown')
                initials = ce.get(qn('w:initials'), '')
                date_str = ce.get(qn('w:date'), '')
                text_elements = ce.xpath('.//w:t')
                text = ''.join(elem.text or '' for elem in text_elements)
                comments.append({
                    'id': f'comment_{idx + 1}',
                    'comment_id': cid,
                    'author': author,
                    'initials': initials,
                    'date': date_str,
                    'text': text.strip(),
                })
    except Exception:
        pass
    return comments


# ---------------------------------------------------------------------------
# Footnote / Endnote helpers (adapted from GongRzhe/Office-Word-MCP-Server)
# ---------------------------------------------------------------------------

def _add_footnote_to_paragraph(paragraph, footnote_text: str):
    """Add a footnote reference at the end of a paragraph with proper XML."""
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "FootnoteReference")
    rPr.append(rStyle)
    run.append(rPr)

    footnoteRef = OxmlElement("w:footnoteReference")
    # Generate a footnote ID
    fn_id = _get_next_footnote_id(paragraph)
    footnoteRef.set(qn("w:id"), str(fn_id))
    run.append(footnoteRef)

    paragraph._p.append(run)

    # Add the footnote content to the footnotes part
    _add_footnote_content(paragraph, fn_id, footnote_text)
    return fn_id


def _get_next_footnote_id(paragraph):
    """Get the next available footnote ID by scanning existing footnotes."""
    try:
        doc_part = paragraph.part
        for rel in doc_part.rels.values():
            if 'footnotes' in rel.reltype:
                footnotes_part = rel.target_part
                existing = footnotes_part.element.findall(qn('w:footnote'))
                max_id = max(int(f.get(qn('w:id'), 0)) for f in existing)
                return max_id + 1
    except Exception:
        pass
    return 1


def _add_footnote_content(paragraph, fn_id: int, text: str):
    """Add footnote content to the document's footnotes part."""
    try:
        doc_part = paragraph.part
        footnotes_part = None
        for rel in doc_part.rels.values():
            if 'footnotes' in rel.reltype:
                footnotes_part = rel.target_part
                break
        if footnotes_part is None:
            return

        fn_xml = (
            f'<w:footnote xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'w:id="{fn_id}" w:type="footnote">'
            f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'
            f'</w:footnote>'
        )
        fn_el = etree.fromstring(fn_xml.encode())
        footnotes_part.element.append(fn_el)
    except Exception:
        pass


def _read_footnotes(doc: Document) -> list[dict]:
    """Read all footnotes from the document."""
    footnotes = []
    try:
        for rel in doc.part.rels.values():
            if 'footnotes' in rel.reltype:
                footnotes_part = rel.target_part
                for fn in footnotes_part.element.findall(qn('w:footnote')):
                    fn_id = fn.get(qn('w:id'), '')
                    fn_type = fn.get(qn('w:type'), 'footnote')
                    text_elements = fn.xpath('.//w:t')
                    text = ''.join(elem.text or '' for elem in text_elements)
                    if fn_type not in ('separator', 'continuationSeparator'):
                        footnotes.append({
                            'id': fn_id,
                            'type': fn_type,
                            'text': text.strip(),
                        })
    except Exception:
        pass
    return footnotes


# ---------------------------------------------------------------------------
# Text highlighting helper (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

def _highlight_text_in_paragraph(paragraph, search: str, color: str = "yellow",
                                  match_case: bool = False) -> int:
    """Highlight occurrences of search text in a paragraph. Returns count."""
    color_map = {
        'yellow': WD_COLOR_INDEX.YELLOW,
        'red': WD_COLOR_INDEX.RED,
        'green': WD_COLOR_INDEX.GREEN,
        'blue': WD_COLOR_INDEX.BLUE,
        'cyan': WD_COLOR_INDEX.CYAN,
        'magenta': WD_COLOR_INDEX.MAGENTA,
        'dark_yellow': WD_COLOR_INDEX.DARK_YELLOW,
        'turquoise': WD_COLOR_INDEX.TURQUOISE,
        'pink': WD_COLOR_INDEX.PINK,
    }
    highlight = color_map.get(color.lower(), WD_COLOR_INDEX.YELLOW)

    full_text = paragraph.text
    search_text = search if match_case else search.lower()
    compare_text = full_text if match_case else full_text.lower()

    if search_text not in compare_text:
        return 0

    # Find all match positions
    matches = []
    pos = 0
    while True:
        idx = compare_text.find(search_text, pos)
        if idx == -1:
            break
        matches.append((idx, idx + len(search_text)))
        pos = idx + len(search_text)

    if not matches:
        return 0

    # Rebuild runs with highlighting
    # Collect all runs with their text and formatting
    runs_info = []
    for run in paragraph.runs:
        runs_info.append({
            'text': run.text,
            'bold': run.bold,
            'italic': run.italic,
            'underline': run.underline,
            'font_name': run.font.name,
            'font_size': run.font.size,
            'font_color': run.font.color.rgb if run.font.color and run.font.color.rgb else None,
        })

    # Clear existing runs
    for run in paragraph.runs:
        run.text = ""

    # Remove all existing run elements
    for run in list(paragraph.runs):
        run._r.getparent().remove(run._r)

    # Rebuild: split text into highlighted and non-highlighted segments
    segments = []
    prev_end = 0
    for start, end in matches:
        if start > prev_end:
            segments.append((full_text[prev_end:start], False))
        segments.append((full_text[start:end], True))
        prev_end = end
    if prev_end < len(full_text):
        segments.append((full_text[prev_end:], False))

    for seg_text, should_highlight in segments:
        if not seg_text:
            continue
        run = paragraph.add_run(seg_text)
        if should_highlight:
            run.font.highlight_color = highlight

    return len(matches)


# ---------------------------------------------------------------------------
# Document protection helpers (adapted from GongRzhe/Office-Word-MCP-Server)
# ---------------------------------------------------------------------------

def _protect_document_with_password(path: str, password: str) -> str:
    """Encrypt a .docx file with a password using msoffcrypto."""
    try:
        import msoffcrypto
    except ImportError:
        return "ERROR: msoffcrypto-tool not installed. Run: pip install msoffcrypto-tool"

    try:
        with open(path, "rb") as f:
            original_data = f.read()

        file = msoffcrypto.OfficeFile(io.BytesIO(original_data))
        file.load_key(password=password)

        encrypted_data = io.BytesIO()
        file.encrypt(password=password, outfile=encrypted_data)

        with open(path, "wb") as f:
            f.write(encrypted_data.getvalue())

        return f"Document encrypted successfully with password: {path}"
    except Exception as e:
        return f"Failed to encrypt document: {str(e)}"


def _unprotect_document_with_password(path: str, password: str) -> str:
    """Decrypt a password-protected .docx file using msoffcrypto."""
    try:
        import msoffcrypto
    except ImportError:
        return "ERROR: msoffcrypto-tool not installed. Run: pip install msoffcrypto-tool"

    try:
        with open(path, "rb") as f:
            encrypted_data = f.read()

        file = msoffcrypto.OfficeFile(io.BytesIO(encrypted_data))
        file.load_key(password=password)

        decrypted_data = io.BytesIO()
        file.decrypt(outfile=decrypted_data)

        with open(path, "wb") as f:
            f.write(decrypted_data.getvalue())

        return f"Document decrypted successfully: {path}"
    except Exception as e:
        return f"Failed to decrypt document: {str(e)}"


# ---------------------------------------------------------------------------
# PDF conversion helper (adapted from GongRzhe/Office-Word-MCP-Server)
# ---------------------------------------------------------------------------

def _convert_to_pdf(path: str, output_path: Optional[str] = None) -> str:
    """Convert a .docx file to PDF. Cross-platform: uses docx2pdf on Windows,
    LibreOffice on Linux/macOS."""
    if not output_path:
        base, _ = os.path.splitext(path)
        output_path = f"{base}.pdf"
    elif not output_path.lower().endswith('.pdf'):
        output_path = f"{output_path}.pdf"

    output_path = os.path.abspath(output_path)
    output_dir = os.path.dirname(output_path) or '.'
    os.makedirs(output_dir, exist_ok=True)

    system = platform.system()

    if system == "Windows":
        try:
            from docx2pdf import convert
            convert(path, output_path)
            return f"Document converted to PDF: {output_path}"
        except ImportError:
            return "ERROR: docx2pdf not installed. Run: pip install docx2pdf"
        except Exception as e:
            return f"Failed to convert to PDF: {str(e)}"

    elif system in ["Linux", "Darwin"]:
        errors = []
        lo_commands = ["soffice"]
        if system == "Darwin":
            lo_commands = ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
        else:
            lo_commands = ["libreoffice", "soffice"]

        for cmd in lo_commands:
            try:
                cmd_list = [cmd, '--headless', '--convert-to', 'pdf',
                           '--outdir', output_dir, path]
                result = subprocess.run(cmd_list, capture_output=True, text=True,
                                       timeout=60, check=False)
                if result.returncode == 0:
                    base_name = os.path.splitext(os.path.basename(path))[0]
                    created_pdf = os.path.join(output_dir, f"{base_name}.pdf")
                    if os.path.exists(created_pdf):
                        if created_pdf != output_path:
                            shutil.move(created_pdf, output_path)
                        return f"Document converted to PDF via {cmd}: {output_path}"
                    errors.append(f"{cmd} succeeded but output file not found")
                else:
                    errors.append(f"{cmd} failed: {result.stderr.strip()}")
            except FileNotFoundError:
                errors.append(f"Command '{cmd}' not found")
            except Exception as e:
                errors.append(f"Error with {cmd}: {str(e)}")

        # Fallback: docx2pdf
        try:
            from docx2pdf import convert
            convert(path, output_path)
            if os.path.exists(output_path):
                return f"Document converted to PDF via docx2pdf: {output_path}"
        except ImportError:
            pass
        except Exception as e:
            errors.append(f"docx2pdf fallback failed: {str(e)}")

        return ("Failed to convert to PDF. Install LibreOffice or docx2pdf.\n"
                f"Errors: {'; '.join(errors)}")
    else:
        return f"PDF conversion not supported on {system}"


# ---------------------------------------------------------------------------
# Document merge helper (adapted from GongRzhe/Office-Word-MCP-Server)
# ---------------------------------------------------------------------------

def _merge_documents(source_paths: list[str], output_path: str) -> str:
    """Merge multiple .docx files into a single document."""
    if not source_paths:
        return "ERROR: No source documents provided"

    for sp in source_paths:
        if not os.path.exists(sp):
            return f"ERROR: Source file not found: {sp}"

    try:
        merged = Document(source_paths[0])
        for sp in source_paths[1:]:
            # Add a page break between documents
            merged.add_page_break()
            sub_doc = Document(sp)
            for element in sub_doc.element.body:
                # Skip sectPr elements (section properties)
                if element.tag == qn('w:sectPr'):
                    continue
                merged.element.body.append(copy.deepcopy(element))

        merged.save(output_path)
        return f"Merged {len(source_paths)} documents into: {output_path}"
    except Exception as e:
        return f"Failed to merge documents: {str(e)}"


# ---------------------------------------------------------------------------
# List insertion helpers (adapted from GongRzhe/Office-Word-MCP-Server)
# ---------------------------------------------------------------------------

def _insert_bulleted_list(doc: Document, items: list[str], style: str = "List Bullet") -> str:
    """Insert a bulleted list into the document."""
    try:
        for item in items:
            para = doc.add_paragraph(item, style=style)
        return f"Bulleted list with {len(items)} items inserted."
    except Exception as e:
        return f"Failed to insert bulleted list: {str(e)}"


def _insert_numbered_list(doc: Document, items: list[str], style: str = "List Number") -> str:
    """Insert a numbered list into the document."""
    try:
        for item in items:
            para = doc.add_paragraph(item, style=style)
        return f"Numbered list with {len(items)} items inserted."
    except Exception as e:
        return f"Failed to insert numbered list: {str(e)}"


# ---------------------------------------------------------------------------
# Cell padding helper (adapted from GongRzhe/Office-Word-MCP-Server)
# ---------------------------------------------------------------------------

def _set_cell_padding(cell, top: float = None, bottom: float = None,
                      left: float = None, right: float = None):
    """Set cell margins/padding in points (converted to twips: 1pt = 20twips)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")

    padding_map = {"top": top, "bottom": bottom, "left": left, "right": right}
    for side, val in padding_map.items():
        if val is not None:
            mar_el = OxmlElement(f"w:{side}")
            mar_el.set(qn("w:w"), str(int(val * 20)))
            mar_el.set(qn("w:type"), "dxa")
            tcMar.append(mar_el)

    # Remove existing tcMar if present
    existing = tcPr.find(qn("w:tcMar"))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(tcMar)


# ---------------------------------------------------------------------------
# Stable anchors (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


def _generate_para_id(used: set[str]) -> str:
    """Generate a valid, unique w14:paraId (8 hex chars, 0x00000001–0x7FFFFFFF)."""
    import random as _random
    for _ in range(10000):
        n = _random.randint(1, 0x7FFFFFFF)
        hid = f"{n:08X}"
        if hid not in used:
            used.add(hid)
            return hid
    raise RuntimeError("Unable to allocate a unique paragraph id.")


def _collect_all_para_ids(doc: Document) -> set[str]:
    """Collect all w14:paraId values from the document body (including tables)."""
    used: set[str] = set()
    body = doc.element.body
    for p_el in body.findall('.//' + qn('w:p')):
        pid = p_el.get(qn('w14:paraId'))
        if pid:
            used.add(pid.upper())
    return used


def _ensure_w14_namespace(doc: Document):
    """Ensure the document root has mc:Ignorable including w14.

    lxml auto-declares xmlns:w14 when a w14:paraId attribute is first set on an
    element, so we only need to update mc:Ignorable here.
    """
    elem = doc.element
    mc_ignorable = elem.get(qn('mc:Ignorable'), '')
    tokens = mc_ignorable.split() if mc_ignorable else []
    if 'w14' not in tokens:
        tokens.append('w14')
        elem.set(qn('mc:Ignorable'), ' '.join(tokens))


def _ensure_anchors(doc: Document) -> dict:
    """Assign stable w14:paraId anchors to every top-level paragraph that lacks one."""
    body = doc.element.body
    used = _collect_all_para_ids(doc)
    _ensure_w14_namespace(doc)

    seeded = 0
    repaired = 0
    blocks = []

    # Track assigned to detect duplicates
    assigned: set[str] = set()

    block_index = 0
    for child in body:
        tag = child.tag
        if tag == qn('w:p'):
            # Get existing paraId
            pid = child.get(qn('w14:paraId'))
            canonical = pid.upper() if pid else None

            needs_reseed = (
                not pid or
                not re.match(r'^[0-9A-Fa-f]{8}$', pid) or
                int(pid, 16) <= 0 or
                int(pid, 16) >= 0x80000000 or
                canonical in assigned
            )

            if not needs_reseed:
                assigned.add(canonical)
            else:
                # Remove stale id
                if pid:
                    del child.attrib[qn('w14:paraId')]
                    repaired += 1
                else:
                    seeded += 1
                fresh = _generate_para_id(used)
                child.set(qn('w14:paraId'), fresh)
                assigned.add(fresh)

            # Get text preview
            text_els = child.findall('.//' + qn('w:t'))
            text = ''.join(t.text or '' for t in text_els).strip()
            preview = text[:50] + '…' if len(text) > 50 else text

            blocks.append({
                'index': block_index,
                'type': 'paragraph',
                'anchor': child.get(qn('w14:paraId')),
                'text_preview': preview,
            })
            block_index += 1
        elif tag == qn('w:tbl'):
            # Tables advance block index but are not anchorable
            text_els = child.findall('.//' + qn('w:t'))
            text = ''.join(t.text or '' for t in text_els).strip()
            preview = text[:50] + '…' if len(text) > 50 else text
            blocks.append({
                'index': block_index,
                'type': 'table',
                'anchor': None,
                'text_preview': preview,
            })
            block_index += 1
        elif tag == qn('w:sdt'):
            text_els = child.findall('.//' + qn('w:t'))
            text = ''.join(t.text or '' for t in text_els).strip()
            preview = text[:50] + '…' if len(text) > 50 else text
            blocks.append({
                'index': block_index,
                'type': 'sdt',
                'anchor': None,
                'text_preview': preview,
            })
            block_index += 1

    return {
        'seeded': seeded,
        'repaired': repaired,
        'blocks': blocks,
    }


# ---------------------------------------------------------------------------
# List images (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

def _list_images(doc: Document) -> list[dict]:
    """List all images embedded in the document."""
    images = []
    try:
        # Access the document's inline shapes
        for i, shape in enumerate(doc.inline_shapes):
            images.append({
                'index': i,
                'type': 'inline',
                'width': shape.width,
                'height': shape.height,
                'width_inches': round(shape.width / 914400, 2) if shape.width else None,
                'height_inches': round(shape.height / 914400, 2) if shape.height else None,
            })
    except Exception:
        pass

    # Also scan for drawing elements in the XML
    try:
        body = doc.element.body
        drawings = body.findall('.//' + qn('w:drawing'))
        for i, drawing in enumerate(drawings):
            # Try to get image relationship
            blip = drawing.find('.//' + qn('a:blip'))
            if blip is not None:
                embed = blip.get(qn('r:embed'))
                if embed:
                    try:
                        rel = doc.part.rels[embed]
                        images.append({
                            'index': len(images),
                            'type': 'drawing',
                            'relationship_id': embed,
                            'target': rel.target_ref if hasattr(rel, 'target_ref') else str(rel._target),
                        })
                    except Exception:
                        pass
    except Exception:
        pass

    return images


# ---------------------------------------------------------------------------
# Threaded comment reply (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

def _reply_to_comment(doc: Document, parent_comment_id: int, reply_text: str,
                      author: str = "Claude") -> str:
    """Reply to an existing comment, creating a threaded conversation."""
    try:
        document_part = doc.part
        comments_part = None
        for rel_id, rel in document_part.rels.items():
            if 'comments' in rel.reltype and 'comments' == rel.reltype.split('/')[-1]:
                comments_part = rel.target_part
                break

        if not comments_part:
            return "ERROR: No comments found in document."

        comments_element = comments_part.element
        # Find parent comment by id
        parent_comment = None
        for ce in comments_element.xpath('.//w:comment'):
            cid = ce.get(qn('w:id'))
            if cid == str(parent_comment_id):
                parent_comment = ce
                break

        if not parent_comment:
            return f"ERROR: Comment with id {parent_comment_id} not found."

        # Append reply paragraph to the parent comment
        reply_para = OxmlElement('w:p')
        reply_run = OxmlElement('w:r')
        reply_rPr = OxmlElement('w:rPr')
        reply_rStyle = OxmlElement('w:rStyle')
        reply_rStyle.set(qn('w:val'), 'CommentReference')
        reply_rPr.append(reply_rStyle)
        reply_run.append(reply_rPr)
        reply_text_el = OxmlElement('w:t')
        reply_text_el.text = reply_text
        reply_run.append(reply_text_el)
        reply_para.append(reply_run)

        # Set author on the paragraph
        reply_para.set(qn('w:author'), author)
        reply_para.set(qn('w:initials'), author[:2])
        reply_para.set(qn('w:date'), datetime.now().isoformat())

        parent_comment.append(reply_para)
        return f"Reply added to comment {parent_comment_id} by {author}."

    except Exception as e:
        return f"ERROR replying to comment: {e}"


# ---------------------------------------------------------------------------
# Delete comment (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

def _delete_comment(doc: Document, comment_id: int) -> str:
    """Delete a comment by its ID and remove range markers from the document."""
    try:
        document_part = doc.part
        comments_part = None
        for rel_id, rel in document_part.rels.items():
            if 'comments' in rel.reltype and 'comments' == rel.reltype.split('/')[-1]:
                comments_part = rel.target_part
                break

        if not comments_part:
            return "ERROR: No comments found in document."

        comments_element = comments_part.element
        # Find and remove the comment
        for ce in comments_element.xpath('.//w:comment'):
            cid = ce.get(qn('w:id'))
            if cid == str(comment_id):
                parent = ce.getparent()
                parent.remove(ce)
                break
        else:
            return f"ERROR: Comment with id {comment_id} not found."

        # Remove comment range markers from document body
        body = doc.element.body
        for tag_name in ['commentRangeStart', 'commentRangeEnd']:
            for marker in body.findall('.//' + qn(f'w:{tag_name}')):
                mid = marker.get(qn('w:id'))
                if mid == str(comment_id):
                    parent = marker.getparent()
                    parent.remove(marker)

        # Remove comment references
        for ref in body.findall('.//' + qn('w:commentReference')):
            rid = ref.get(qn('w:id'))
            if rid == str(comment_id):
                parent = ref.getparent()
                parent.remove(ref)

        return f"Comment {comment_id} deleted successfully."

    except Exception as e:
        return f"ERROR deleting comment: {e}"


# ---------------------------------------------------------------------------
# Get page layout (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

def _get_page_layout(doc: Document) -> dict:
    """Get page size, margins, and orientation of the document."""
    section = doc.sections[0]
    layout = {
        'page_width': section.page_width,
        'page_height': section.page_height,
        'page_width_mm': round(section.page_width / 36000, 1) if section.page_width else None,
        'page_height_mm': round(section.page_height / 36000, 1) if section.page_height else None,
        'orientation': 'landscape' if section.orientation == WD_ORIENT.LANDSCAPE else 'portrait',
        'top_margin': section.top_margin,
        'bottom_margin': section.bottom_margin,
        'left_margin': section.left_margin,
        'right_margin': section.right_margin,
        'top_margin_mm': round(section.top_margin / 36000, 1) if section.top_margin else None,
        'bottom_margin_mm': round(section.bottom_margin / 36000, 1) if section.bottom_margin else None,
        'left_margin_mm': round(section.left_margin / 36000, 1) if section.left_margin else None,
        'right_margin_mm': round(section.right_margin / 36000, 1) if section.right_margin else None,
        'header_distance': section.header_distance,
        'footer_distance': section.footer_distance,
        'gutter': section.gutter,
    }

    # Detect preset
    w_mm = layout['page_width_mm']
    h_mm = layout['page_height_mm']
    presets = {
        'A4': (210, 297), 'A3': (297, 420), 'A5': (148, 210),
        'LETTER': (215.9, 279.4), 'LEGAL': (215.9, 355.6),
        'B4': (250, 353), 'B5': (176, 250),
    }
    detected = None
    for name, (pw, ph) in presets.items():
        if abs(w_mm - pw) < 2 and abs(h_mm - ph) < 2:
            detected = name
            break
    layout['page_size_preset'] = detected

    return layout


# ---------------------------------------------------------------------------
# Search-and-format text (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

def _format_text_in_paragraphs(doc: Document, search: str,
                                bold: bool = None, italic: bool = None,
                                underline: bool = None, strikethrough: bool = None,
                                highlight_color: str = None,
                                font_name: str = None, font_size: float = None,
                                font_color: str = None,
                                match_case: bool = False) -> int:
    """Apply character formatting to all runs matching the search text."""
    count = 0
    search_text = search if match_case else search.lower()

    color_map = {
        'yellow': WD_COLOR_INDEX.YELLOW, 'green': WD_COLOR_INDEX.GREEN,
        'cyan': WD_COLOR_INDEX.CYAN, 'magenta': WD_COLOR_INDEX.MAGENTA,
        'blue': WD_COLOR_INDEX.BLUE, 'red': WD_COLOR_INDEX.RED,
        'darkblue': WD_COLOR_INDEX.DARK_BLUE, 'darkcyan': WD_COLOR_INDEX.DARK_CYAN,
        'darkgreen': WD_COLOR_INDEX.DARK_GREEN, 'darkmagenta': WD_COLOR_INDEX.DARK_MAGENTA,
        'darkred': WD_COLOR_INDEX.DARK_RED, 'darkyellow': WD_COLOR_INDEX.DARK_YELLOW,
        'darkgray': WD_COLOR_INDEX.DARK_GRAY, 'lightgray': WD_COLOR_INDEX.LIGHT_GRAY,
        'black': WD_COLOR_INDEX.BLACK, 'none': WD_COLOR_INDEX.AUTO,
    }

    def _format_run(run):
        if bold is not None:
            run.bold = bold
        if italic is not None:
            run.italic = italic
        if underline is not None:
            run.underline = underline
        if strikethrough is not None:
            run.font.strike = strikethrough
        if highlight_color:
            hl = color_map.get(highlight_color.lower(), WD_COLOR_INDEX.YELLOW)
            run.font.highlight_color = hl
        if font_name:
            run.font.name = font_name
        if font_size:
            run.font.size = Pt(font_size)
        if font_color:
            h = font_color.lstrip('#')
            run.font.color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    def _process_paragraph(para):
        nonlocal count
        full_text = para.text
        compare_text = full_text if match_case else full_text.lower()
        idx = 0
        while True:
            pos = compare_text.find(search_text, idx)
            if pos == -1:
                break
            # Find and format the runs that contain this range
            char_pos = 0
            for run in para.runs:
                run_len = len(run.text)
                run_start = char_pos
                run_end = char_pos + run_len
                # Check if this run overlaps with the match
                if run_start < pos + len(search_text) and run_end > pos:
                    _format_run(run)
                    count += 1
                char_pos = run_end
            idx = pos + len(search_text)

    for para in doc.paragraphs:
        _process_paragraph(para)

    # Also process table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _process_paragraph(para)

    return count


# ---------------------------------------------------------------------------
# Bulk set headings (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

def _set_headings_bulk(doc: Document, headings: list[dict]) -> str:
    """Convert multiple paragraphs to headings in one operation."""
    results = []
    paras = doc.paragraphs
    for h in headings:
        para_idx = h.get('paragraph_index')
        level = h.get('level', 1)
        if para_idx is None or para_idx < 1 or para_idx > len(paras):
            results.append(f"  SKIP: Invalid paragraph_index {para_idx}")
            continue
        para = paras[para_idx - 1]
        style_name = f"Heading {level}"
        try:
            para.style = doc.styles[style_name]
            results.append(f"  Paragraph {para_idx} -> {style_name}")
        except KeyError:
            results.append(f"  ERROR: Style '{style_name}' not found in document.")
    return f"Set {len(headings)} headings:\n" + '\n'.join(results)


# ---------------------------------------------------------------------------
# Get paragraph format (adapted from knorq-ai/docx-mcp-server)
# ---------------------------------------------------------------------------

def _get_paragraph_format(doc: Document, para_idx: int) -> dict:
    """Introspect a paragraph's formatting."""
    paras = doc.paragraphs
    if para_idx < 1 or para_idx > len(paras):
        return {'error': f'Invalid paragraph index. Document has {len(paras)} paragraphs (1-{len(paras)}).'}
    para = paras[para_idx - 1]
    fmt = para.paragraph_format
    info = {
        'paragraph_index': para_idx,
        'style': para.style.name if para.style else None,
        'text_preview': para.text[:100],
        'alignment': str(para.alignment) if para.alignment else None,
        'space_before': fmt.space_before,
        'space_after': fmt.space_after,
        'line_spacing': fmt.line_spacing,
        'left_indent': fmt.left_indent,
        'right_indent': fmt.right_indent,
        'first_line_indent': fmt.first_line_indent,
    }

    # Check for heading level
    if para.style and para.style.name.startswith('Heading'):
        try:
            info['heading_level'] = int(para.style.name.split()[-1])
        except (ValueError, IndexError):
            pass

    # Run formatting info
    runs_info = []
    for run in para.runs:
        ri = {
            'text': run.text[:50],
            'bold': run.bold,
            'italic': run.italic,
            'underline': run.underline,
            'font_name': run.font.name,
            'font_size': run.font.size,
        }
        if run.font.color and run.font.color.rgb:
            ri['font_color'] = str(run.font.color.rgb)
        runs_info.append(ri)
    info['runs'] = runs_info

    return info


# ---------------------------------------------------------------------------
# Safe-call wrapper with fallback (self-recovery)
# ---------------------------------------------------------------------------

def _safe_call(func, *args, **kwargs):
    """Call a function with fallback error handling and recovery suggestions."""
    try:
        return func(*args, **kwargs)
    except Exception as e:
        err_msg = str(e)
        # Provide helpful fallback suggestions
        suggestions = []
        if "No document open" in err_msg:
            suggestions.append("Try: open_document(path='...') or create_new_document(path='...')")
        elif "list index out of range" in err_msg.lower() or "index" in err_msg.lower():
            suggestions.append("Check: use get_document_info or read_document to verify valid indices.")
        elif "style" in err_msg.lower() and "not found" in err_msg.lower():
            suggestions.append("Try: list_styles to see available styles, or create_style to make a new one.")
        elif "file" in err_msg.lower() and ("not found" in err_msg.lower() or "no such" in err_msg.lower()):
            suggestions.append("Check: verify the file path exists and is a valid .docx file.")
        elif "permission" in err_msg.lower():
            suggestions.append("Check: ensure the file is not open in another program (e.g. Microsoft Word).")
        elif "msoffcrypto" in err_msg.lower():
            suggestions.append("Install: pip install msoffcrypto-tool")
        elif "docx2pdf" in err_msg.lower():
            suggestions.append("Install: pip install docx2pdf (Windows) or LibreOffice (Linux/macOS)")

        fallback_msg = f"ERROR: {e}"
        if suggestions:
            fallback_msg += "\n\nFallback suggestions:\n" + '\n'.join(f"  - {s}" for s in suggestions)
        return fallback_msg


# ---------------------------------------------------------------------------
# Tool definitions
# ---------------------------------------------------------------------------

@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        # ---- Document lifecycle ----
        Tool(
            name="open_document",
            description="Open an existing .docx file for editing.",
            inputSchema={
                "type": "object",
                "required": ["path"],
                "properties": {
                    "path": {"type": "string", "description": "Absolute path to the .docx file"}
                }
            }
        ),
        Tool(
            name="create_new_document",
            description="Create a new blank .docx file, optionally from a template.",
            inputSchema={
                "type": "object",
                "required": ["path"],
                "properties": {
                    "path": {"type": "string"},
                    "template": {"type": "string", "description": "Path to a .docx template file"},
                    "default_font": {"type": "string", "description": "Default font name for Normal style"},
                    "default_font_size": {"type": "number", "description": "Default font size in points"}
                }
            }
        ),
        Tool(
            name="save_document",
            description="Save the current document (overwrites original). Creates a .bak backup by default.",
            inputSchema={
                "type": "object",
                "properties": {
                    "backup": {"type": "boolean", "description": "Create .bak backup before saving (default true)"}
                }
            }
        ),
        Tool(
            name="save_as",
            description="Save the current document to a new file path.",
            inputSchema={
                "type": "object",
                "required": ["path"],
                "properties": {"path": {"type": "string"}}
            }
        ),
        Tool(
            name="close_document",
            description="Close the current document without saving.",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="duplicate_document",
            description="Copy the current document file to a new path.",
            inputSchema={
                "type": "object",
                "required": ["new_path"],
                "properties": {"new_path": {"type": "string"}}
            }
        ),

        # ---- Read & inspect ----
        Tool(
            name="get_document_info",
            description="Get document metadata: title, author, dates, word count, paragraph count, sections.",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="read_document",
            description=(
                "Read the full document content as structured JSON. "
                "Returns paragraphs (index, style, text, runs) and tables."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "include_xml": {"type": "boolean", "description": "Include raw XML for each paragraph (default false)"},
                    "include_tables": {"type": "boolean", "description": "Include tables (default true)"}
                }
            }
        ),
        Tool(
            name="get_outline",
            description="Get the document heading hierarchy (outline/TOC preview).",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="get_section",
            description="Get all content paragraphs under a specific heading.",
            inputSchema={
                "type": "object",
                "required": ["heading"],
                "properties": {
                    "heading": {"type": "string", "description": "Heading text (partial match OK)"}
                }
            }
        ),
        Tool(
            name="find_text",
            description="Search for text in the document. Returns all matching paragraphs with index and context.",
            inputSchema={
                "type": "object",
                "required": ["query"],
                "properties": {
                    "query": {"type": "string"},
                    "case_sensitive": {"type": "boolean", "description": "Default false"},
                    "include_tables": {"type": "boolean", "description": "Also search table cells (default true)"}
                }
            }
        ),
        Tool(
            name="read_table",
            description="Read one or all tables in the document.",
            inputSchema={
                "type": "object",
                "properties": {
                    "table_index": {"type": "integer", "description": "Omit to read all tables"}
                }
            }
        ),
        Tool(
            name="list_styles",
            description="List all available styles in the document.",
            inputSchema={
                "type": "object",
                "properties": {
                    "style_type": {
                        "type": "string",
                        "enum": ["paragraph", "character", "table", "all"],
                        "description": "Default: paragraph"
                    }
                }
            }
        ),
        Tool(
            name="get_document_xml",
            description="Get the raw OOXML of a specific paragraph or the entire document body.",
            inputSchema={
                "type": "object",
                "properties": {
                    "paragraph_index": {"type": "integer"},
                    "full_document": {"type": "boolean"}
                }
            }
        ),
        Tool(
            name="read_headers_footers",
            description="Read the content of headers and footers for all sections.",
            inputSchema={"type": "object", "properties": {}}
        ),

        # ---- Text editing ----
        Tool(
            name="replace_text",
            description=(
                "Find and replace text throughout the document. "
                "Preserves run formatting. Searches paragraphs and optionally table cells."
            ),
            inputSchema={
                "type": "object",
                "required": ["find", "replace"],
                "properties": {
                    "find": {"type": "string"},
                    "replace": {"type": "string"},
                    "case_sensitive": {"type": "boolean"},
                    "whole_word": {"type": "boolean"},
                    "include_tables": {"type": "boolean", "description": "Also replace in table cells (default true)"}
                }
            }
        ),
        Tool(
            name="replace_text_batch",
            description=(
                "Batch version of replace_text: perform multiple find-and-replace operations in one call. "
                "Each item in the replacements array is a {find, replace} object. "
                "Useful for applying many text changes at once (e.g. fixing multiple terms, dates, names). "
                "Replacements are applied in order. Saves once at the end."
            ),
            inputSchema={
                "type": "object",
                "required": ["replacements"],
                "properties": {
                    "replacements": {
                        "type": "array",
                        "description": "Array of find-and-replace objects",
                        "items": {
                            "type": "object",
                            "required": ["find", "replace"],
                            "properties": {
                                "find": {"type": "string", "description": "Text to find"},
                                "replace": {"type": "string", "description": "Text to replace with"},
                                "case_sensitive": {"type": "boolean", "description": "Case-sensitive match (default false)"},
                                "whole_word": {"type": "boolean", "description": "Whole word match (default false)"},
                                "include_tables": {"type": "boolean", "description": "Also replace in table cells (default true)"}
                            }
                        }
                    }
                }
            }
        ),
        Tool(
            name="replace_paragraph",
            description="Replace the full text content of a specific paragraph (by index or text match).",
            inputSchema={
                "type": "object",
                "required": ["new_text"],
                "properties": {
                    "match_text": {"type": "string", "description": "Partial text to identify the paragraph"},
                    "paragraph_index": {"type": "integer"},
                    "new_text": {"type": "string"},
                    "preserve_style": {"type": "boolean", "description": "Keep original style (default true)"}
                }
            }
        ),
        Tool(
            name="replace_section",
            description="Replace all content paragraphs under a heading with new content.",
            inputSchema={
                "type": "object",
                "required": ["heading", "new_paragraphs"],
                "properties": {
                    "heading": {"type": "string"},
                    "new_paragraphs": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "text": {"type": "string"},
                                "style": {"type": "string"}
                            }
                        }
                    }
                }
            }
        ),
        Tool(
            name="insert_paragraph",
            description=(
                "Insert a new paragraph at a given position. "
                "Position can be: after_text, after_heading, at_index, or at_end."
            ),
            inputSchema={
                "type": "object",
                "required": ["text"],
                "properties": {
                    "text": {"type": "string"},
                    "style": {"type": "string", "description": "Style name e.g. 'Normal', 'Heading 1', 'List Bullet'"},
                    "after_text": {"type": "string"},
                    "after_heading": {"type": "string"},
                    "at_index": {"type": "integer"},
                    "at_end": {"type": "boolean"},
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "font_size": {"type": "number"},
                    "font_name": {"type": "string"},
                    "color_hex": {"type": "string"},
                    "alignment": {"type": "string", "enum": ["LEFT", "CENTER", "RIGHT", "JUSTIFY"]}
                }
            }
        ),
        Tool(
            name="delete_paragraph",
            description="Delete one or more paragraphs by index or text match.",
            inputSchema={
                "type": "object",
                "properties": {
                    "match_text": {"type": "string"},
                    "paragraph_index": {"type": "integer"},
                    "delete_all_matching": {"type": "boolean", "description": "Delete all matches (default false, only first)"}
                }
            }
        ),
        Tool(
            name="delete_section",
            description="Delete a heading paragraph and all content under it.",
            inputSchema={
                "type": "object",
                "required": ["heading"],
                "properties": {"heading": {"type": "string"}}
            }
        ),
        Tool(
            name="move_section",
            description="Move a section (heading + its content) to before or after another heading.",
            inputSchema={
                "type": "object",
                "required": ["section_heading"],
                "properties": {
                    "section_heading": {"type": "string"},
                    "before_heading": {"type": "string"},
                    "after_heading": {"type": "string"}
                }
            }
        ),

        # ---- Formatting ----
        Tool(
            name="format_paragraph",
            description="Change paragraph-level formatting: style, alignment, spacing, indentation.",
            inputSchema={
                "type": "object",
                "properties": {
                    "match_text": {"type": "string"},
                    "paragraph_index": {"type": "integer"},
                    "style": {"type": "string"},
                    "alignment": {"type": "string", "enum": ["LEFT", "CENTER", "RIGHT", "JUSTIFY"]},
                    "space_before": {"type": "number", "description": "Points before paragraph"},
                    "space_after": {"type": "number", "description": "Points after paragraph"},
                    "left_indent": {"type": "number", "description": "Left indent in cm"},
                    "right_indent": {"type": "number", "description": "Right indent in cm"},
                    "first_line_indent": {"type": "number", "description": "First-line indent in cm"},
                    "line_spacing": {"type": "number", "description": "Line spacing in points (e.g. 24 = double)"},
                    "keep_together": {"type": "boolean"},
                    "keep_with_next": {"type": "boolean"},
                    "page_break_before": {"type": "boolean"}
                }
            }
        ),
        Tool(
            name="format_text_run",
            description=(
                "Apply character-level formatting to runs within a paragraph. "
                "Can target a specific run by text match."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "paragraph_match": {"type": "string", "description": "Text to find the paragraph"},
                    "paragraph_index": {"type": "integer"},
                    "run_text_match": {"type": "string", "description": "Narrow to runs containing this text"},
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "underline": {"type": "boolean"},
                    "strike": {"type": "boolean"},
                    "font_size": {"type": "number"},
                    "font_name": {"type": "string"},
                    "color_hex": {"type": "string", "description": "6-char hex e.g. 'FF0000'"},
                    "all_caps": {"type": "boolean"},
                    "small_caps": {"type": "boolean"}
                }
            }
        ),

        # ---- Headings ----
        Tool(
            name="add_heading",
            description="Add a new heading at the end of the document or after a specific heading.",
            inputSchema={
                "type": "object",
                "required": ["text", "level"],
                "properties": {
                    "text": {"type": "string"},
                    "level": {"type": "integer", "description": "Heading level 1-9"},
                    "after_heading": {"type": "string", "description": "Insert after this heading"},
                    "at_end": {"type": "boolean", "description": "Append at end (default true)"}
                }
            }
        ),

        # ---- Tables ----
        Tool(
            name="insert_table",
            description=(
                "Insert a table with optional data, style, and column widths. "
                "Supports header row formatting and cell-level formatting via dict values."
            ),
            inputSchema={
                "type": "object",
                "required": ["rows", "cols"],
                "properties": {
                    "rows": {"type": "integer"},
                    "cols": {"type": "integer"},
                    "data": {
                        "type": "array",
                        "description": "Array of rows; each row is an array of strings or {text, bold, color_hex} dicts",
                        "items": {"type": "array"}
                    },
                    "style": {"type": "string", "description": "Table style name (default: 'Table Grid')"},
                    "header_row": {"type": "boolean", "description": "Make first row bold"},
                    "col_widths": {"type": "array", "items": {"type": "number"}, "description": "Column widths in cm"},
                    "after_text": {"type": "string", "description": "Insert after paragraph containing this text"},
                    "after_heading": {"type": "string"}
                }
            }
        ),
        Tool(
            name="edit_table_cell",
            description="Edit the text of a specific table cell.",
            inputSchema={
                "type": "object",
                "required": ["table_index", "row", "col", "text"],
                "properties": {
                    "table_index": {"type": "integer"},
                    "row": {"type": "integer"},
                    "col": {"type": "integer"},
                    "text": {"type": "string"},
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "font_size": {"type": "number"},
                    "color_hex": {"type": "string"},
                    "bg_color": {"type": "string", "description": "Cell background color hex (no #)"},
                    "alignment": {"type": "string", "enum": ["LEFT", "CENTER", "RIGHT", "JUSTIFY"]}
                }
            }
        ),
        Tool(
            name="add_table_row",
            description="Append a new row to an existing table.",
            inputSchema={
                "type": "object",
                "required": ["table_index"],
                "properties": {
                    "table_index": {"type": "integer"},
                    "data": {"type": "array", "items": {"type": "string"}}
                }
            }
        ),
        Tool(
            name="merge_table_cells",
            description="Merge table cells horizontally or vertically.",
            inputSchema={
                "type": "object",
                "required": ["table_index", "direction"],
                "properties": {
                    "table_index": {"type": "integer"},
                    "direction": {"type": "string", "enum": ["horizontal", "vertical"]},
                    "row": {"type": "integer", "description": "Row index (for horizontal merge)"},
                    "col": {"type": "integer", "description": "Column index (for vertical merge)"},
                    "start_col": {"type": "integer"},
                    "end_col": {"type": "integer"},
                    "start_row": {"type": "integer"},
                    "end_row": {"type": "integer"}
                }
            }
        ),
        Tool(
            name="format_table_cell",
            description="Apply background color and/or borders to a table cell.",
            inputSchema={
                "type": "object",
                "required": ["table_index", "row", "col"],
                "properties": {
                    "table_index": {"type": "integer"},
                    "row": {"type": "integer"},
                    "col": {"type": "integer"},
                    "bg_color": {"type": "string", "description": "Fill color hex e.g. 'FFCC00'"},
                    "border_top": {"type": "string", "enum": ["single", "double", "none"]},
                    "border_bottom": {"type": "string", "enum": ["single", "double", "none"]},
                    "border_left": {"type": "string", "enum": ["single", "double", "none"]},
                    "border_right": {"type": "string", "enum": ["single", "double", "none"]},
                    "border_color": {"type": "string", "description": "Border color hex (default 000000)"},
                    "border_size": {"type": "integer", "description": "Border size in eighths of a point (default 4)"}
                }
            }
        ),
        Tool(
            name="delete_table_row",
            description="Delete a row from a table.",
            inputSchema={
                "type": "object",
                "required": ["table_index", "row"],
                "properties": {
                    "table_index": {"type": "integer"},
                    "row": {"type": "integer"}
                }
            }
        ),
        Tool(
            name="delete_table",
            description="Delete an entire table from the document.",
            inputSchema={
                "type": "object",
                "required": ["table_index"],
                "properties": {
                    "table_index": {"type": "integer"}
                }
            }
        ),

        # ---- Headers & Footers ----
        Tool(
            name="set_header",
            description=(
                "Set the header text for a document section. "
                "Supports left/center/right tabs, page numbers, and date fields."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "text": {"type": "string", "description": "Header text"},
                    "section_index": {"type": "integer", "description": "Section index (default 0)"},
                    "which": {
                        "type": "string",
                        "enum": ["header", "first_page_header", "even_page_header"],
                        "description": "Default: header"
                    },
                    "alignment": {"type": "string", "enum": ["LEFT", "CENTER", "RIGHT"]},
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "font_size": {"type": "number"},
                    "font_name": {"type": "string"}
                }
            }
        ),
        Tool(
            name="set_footer",
            description=(
                "Set the footer text for a document section. "
                "Supports page numbers and custom text."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "text": {"type": "string", "description": "Footer text (use '' to insert only page number)"},
                    "section_index": {"type": "integer", "description": "Section index (default 0)"},
                    "which": {
                        "type": "string",
                        "enum": ["footer", "first_page_footer", "even_page_footer"],
                        "description": "Default: footer"
                    },
                    "alignment": {"type": "string", "enum": ["LEFT", "CENTER", "RIGHT"]},
                    "add_page_number": {"type": "boolean", "description": "Add PAGE field"},
                    "page_number_format": {
                        "type": "string",
                        "enum": ["PAGE_NUMBER", "PAGE_OF_PAGES"],
                        "description": "Page number format"
                    },
                    "bold": {"type": "boolean"},
                    "font_size": {"type": "number"},
                    "font_name": {"type": "string"}
                }
            }
        ),
        Tool(
            name="add_image_to_header",
            description="Insert an image (e.g. company logo) into a header.",
            inputSchema={
                "type": "object",
                "required": ["image_path"],
                "properties": {
                    "image_path": {"type": "string", "description": "Absolute path to image file (PNG, JPG, etc.)"},
                    "section_index": {"type": "integer", "description": "Default 0"},
                    "width_cm": {"type": "number", "description": "Image width in cm"},
                    "alignment": {"type": "string", "enum": ["LEFT", "CENTER", "RIGHT"]}
                }
            }
        ),
        Tool(
            name="clear_header",
            description="Remove all content from a header.",
            inputSchema={
                "type": "object",
                "properties": {
                    "section_index": {"type": "integer", "description": "Default 0"},
                    "which": {"type": "string", "enum": ["header", "first_page_header", "even_page_header"]}
                }
            }
        ),
        Tool(
            name="clear_footer",
            description="Remove all content from a footer.",
            inputSchema={
                "type": "object",
                "properties": {
                    "section_index": {"type": "integer", "description": "Default 0"},
                    "which": {"type": "string", "enum": ["footer", "first_page_footer", "even_page_footer"]}
                }
            }
        ),

        # ---- Images ----
        Tool(
            name="insert_image",
            description="Insert an image into the document body at a given position.",
            inputSchema={
                "type": "object",
                "required": ["image_path"],
                "properties": {
                    "image_path": {"type": "string", "description": "Absolute path to image file"},
                    "width_cm": {"type": "number", "description": "Width in cm (height auto-scaled)"},
                    "height_cm": {"type": "number", "description": "Height in cm (optional)"},
                    "alignment": {"type": "string", "enum": ["LEFT", "CENTER", "RIGHT"]},
                    "after_text": {"type": "string"},
                    "after_heading": {"type": "string"},
                    "at_end": {"type": "boolean"},
                    "caption": {"type": "string", "description": "Optional caption paragraph below image"}
                }
            }
        ),

        # ---- Page layout ----
        Tool(
            name="set_page_margins",
            description="Set page margins for all or a specific section.",
            inputSchema={
                "type": "object",
                "properties": {
                    "top": {"type": "number", "description": "Top margin in cm"},
                    "bottom": {"type": "number", "description": "Bottom margin in cm"},
                    "left": {"type": "number", "description": "Left margin in cm"},
                    "right": {"type": "number", "description": "Right margin in cm"},
                    "section_index": {"type": "integer", "description": "Omit to apply to all sections"}
                }
            }
        ),
        Tool(
            name="set_page_orientation",
            description="Set page orientation for a section.",
            inputSchema={
                "type": "object",
                "required": ["orientation"],
                "properties": {
                    "orientation": {"type": "string", "enum": ["portrait", "landscape"]},
                    "section_index": {"type": "integer", "description": "Default 0"}
                }
            }
        ),
        Tool(
            name="set_page_size",
            description="Set the page size (e.g. A4, Letter, or custom dimensions).",
            inputSchema={
                "type": "object",
                "properties": {
                    "preset": {
                        "type": "string",
                        "enum": ["A4", "A3", "A5", "Letter", "Legal"],
                        "description": "Use a preset size"
                    },
                    "width_cm": {"type": "number", "description": "Custom width in cm"},
                    "height_cm": {"type": "number", "description": "Custom height in cm"},
                    "section_index": {"type": "integer", "description": "Default 0"}
                }
            }
        ),
        Tool(
            name="add_section_break",
            description="Insert a section break (new page, continuous, even page, odd page).",
            inputSchema={
                "type": "object",
                "properties": {
                    "break_type": {
                        "type": "string",
                        "enum": ["new_page", "continuous", "even_page", "odd_page"],
                        "description": "Default: new_page"
                    },
                    "after_text": {"type": "string"},
                    "at_end": {"type": "boolean"}
                }
            }
        ),
        Tool(
            name="add_page_break",
            description="Insert a hard page break after a specific paragraph.",
            inputSchema={
                "type": "object",
                "properties": {
                    "after_text": {"type": "string"},
                    "at_index": {"type": "integer"},
                    "at_end": {"type": "boolean"}
                }
            }
        ),
        Tool(
            name="set_columns",
            description="Set multi-column layout for a section.",
            inputSchema={
                "type": "object",
                "required": ["num_cols"],
                "properties": {
                    "num_cols": {"type": "integer", "description": "Number of columns (1 to disable)"},
                    "spacing_cm": {"type": "number", "description": "Space between columns in cm (default 1.25)"},
                    "section_index": {"type": "integer", "description": "Default 0"}
                }
            }
        ),

        # ---- Styles ----
        Tool(
            name="create_style",
            description="Create or update a custom paragraph style.",
            inputSchema={
                "type": "object",
                "required": ["style_name"],
                "properties": {
                    "style_name": {"type": "string"},
                    "base_style": {"type": "string", "description": "Parent style (default: Normal)"},
                    "font_name": {"type": "string"},
                    "font_size": {"type": "number"},
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "color_hex": {"type": "string"},
                    "alignment": {"type": "string", "enum": ["LEFT", "CENTER", "RIGHT", "JUSTIFY"]},
                    "space_before": {"type": "number", "description": "Points"},
                    "space_after": {"type": "number", "description": "Points"}
                }
            }
        ),

        # ---- TOC ----
        Tool(
            name="insert_toc",
            description=(
                "Insert a Table of Contents field. "
                "Word will render it on first open (press Ctrl+A then F9 to update)."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "title": {"type": "string", "description": "TOC heading text (default: 'Table of Contents')"},
                    "max_level": {"type": "integer", "description": "Max heading level to include (default 3)"},
                    "after_heading": {"type": "string"},
                    "at_end": {"type": "boolean"}
                }
            }
        ),

        # ---- Hyperlinks ----
        Tool(
            name="add_hyperlink",
            description="Add a hyperlink to an existing paragraph.",
            inputSchema={
                "type": "object",
                "required": ["paragraph_match", "text", "url"],
                "properties": {
                    "paragraph_match": {"type": "string", "description": "Text to identify the target paragraph"},
                    "paragraph_index": {"type": "integer"},
                    "text": {"type": "string", "description": "Display text for the hyperlink"},
                    "url": {"type": "string", "description": "URL for the hyperlink"},
                    "color_hex": {"type": "string", "description": "Link color hex (default: 0563C1)"},
                    "underline": {"type": "boolean", "description": "Default true"}
                }
            }
        ),

        # ---- Comments ----
        Tool(
            name="add_comment",
            description="Add a native Word comment to a paragraph (visible in the review pane).",
            inputSchema={
                "type": "object",
                "required": ["match_text", "comment"],
                "properties": {
                    "match_text": {"type": "string"},
                    "comment": {"type": "string"},
                    "author": {"type": "string", "description": "Comment author (default: Claude)"},
                    "initials": {"type": "string", "description": "Author initials (default: AI)"}
                }
            }
        ),

        # ---- Bookmarks ----
        Tool(
            name="add_bookmark",
            description="Add a named bookmark to a paragraph.",
            inputSchema={
                "type": "object",
                "required": ["match_text", "bookmark_name"],
                "properties": {
                    "match_text": {"type": "string"},
                    "paragraph_index": {"type": "integer"},
                    "bookmark_name": {"type": "string"},
                    "bookmark_id": {"type": "integer", "description": "Unique numeric ID (default: 1)"}
                }
            }
        ),

        # ---- Document properties ----
        Tool(
            name="set_document_properties",
            description="Set document core properties (title, author, subject, keywords, description, category).",
            inputSchema={
                "type": "object",
                "properties": {
                    "title": {"type": "string"},
                    "author": {"type": "string"},
                    "subject": {"type": "string"},
                    "keywords": {"type": "string"},
                    "description": {"type": "string"},
                    "category": {"type": "string"},
                    "company": {"type": "string"}
                }
            }
        ),

        # ---- Advanced XML ----
        Tool(
            name="apply_xml_patch",
            description="Replace a paragraph's XML with custom OOXML (for advanced formatting not covered by other tools).",
            inputSchema={
                "type": "object",
                "required": ["paragraph_index", "xml_content"],
                "properties": {
                    "paragraph_index": {"type": "integer"},
                    "xml_content": {"type": "string", "description": "Full XML of the new w:p element"}
                }
            }
        ),

        # ---- Batch builder ----
        Tool(
            name="build_document",
            description=(
                "Create a fully-formatted document in ONE call from a structured JSON spec. "
                "Supports: headings (H1-H9), paragraphs, bullet/numbered lists, tables, "
                "page breaks, section breaks, images, hyperlinks, per-run formatting "
                "(bold, italic, underline, strikethrough, font, size, color), "
                "header/footer, TOC, and document properties. "
                "Use this for creating new documents with substantial content."
            ),
            inputSchema={
                "type": "object",
                "required": ["path", "elements"],
                "properties": {
                    "path": {"type": "string"},
                    "template": {"type": "string", "description": "Path to .docx template"},
                    "default_font": {"type": "string"},
                    "default_font_size": {"type": "number"},
                    "title": {"type": "string"},
                    "author": {"type": "string"},
                    "subject": {"type": "string"},
                    "header_text": {"type": "string", "description": "Global header for all sections"},
                    "footer_text": {"type": "string"},
                    "footer_page_numbers": {"type": "boolean", "description": "Add page numbers to footer"},
                    "margins": {
                        "type": "object",
                        "properties": {
                            "top": {"type": "number"}, "bottom": {"type": "number"},
                            "left": {"type": "number"}, "right": {"type": "number"}
                        }
                    },
                    "elements": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "type": {
                                    "type": "string",
                                    "enum": ["heading", "paragraph", "list", "table",
                                             "page_break", "section_break", "image",
                                             "toc", "hyperlink"]
                                },
                                "text": {"type": "string"},
                                "level": {"type": "integer"},
                                "style": {"type": "string"},
                                "alignment": {"type": "string", "enum": ["LEFT", "CENTER", "RIGHT", "JUSTIFY"]},
                                "space_before": {"type": "number"},
                                "space_after": {"type": "number"},
                                "runs": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "properties": {
                                            "text": {"type": "string"},
                                            "bold": {"type": "boolean"},
                                            "italic": {"type": "boolean"},
                                            "underline": {"type": "boolean"},
                                            "strike": {"type": "boolean"},
                                            "font_size": {"type": "number"},
                                            "font_name": {"type": "string"},
                                            "color_hex": {"type": "string"},
                                            "url": {"type": "string", "description": "Makes this run a hyperlink"}
                                        }
                                    }
                                },
                                "list_type": {"type": "string", "enum": ["bullet", "numbered"]},
                                "list_level": {"type": "integer"},
                                "data": {"type": "array"},
                                "header_row": {"type": "boolean"},
                                "table_style": {"type": "string"},
                                "col_widths": {"type": "array"},
                                "image_path": {"type": "string"},
                                "width_cm": {"type": "number"},
                                "caption": {"type": "string"},
                                "url": {"type": "string"},
                                "toc_title": {"type": "string"},
                                "toc_max_level": {"type": "integer"},
                                "break_type": {"type": "string"}
                            }
                        }
                    }
                }
            }
        ),

        # ====================================================================
        # Tracked changes (adapted from knorq-ai/docx-mcp-server)
        # ====================================================================
        Tool(
            name="accept_all_changes",
            description="Accept all tracked changes in the current document.",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="reject_all_changes",
            description="Reject all tracked changes in the current document.",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="has_tracked_changes",
            description="Check if the current document contains any tracked changes.",
            inputSchema={"type": "object", "properties": {}}
        ),

        # ====================================================================
        # Comment extraction (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================
        Tool(
            name="get_all_comments",
            description="Extract all comments from the current document with metadata (author, date, text).",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="get_comments_by_author",
            description="Extract comments filtered by author name (case-insensitive).",
            inputSchema={
                "type": "object",
                "required": ["author"],
                "properties": {
                    "author": {"type": "string", "description": "Author name to filter by"}
                }
            }
        ),

        # ====================================================================
        # Footnotes & Endnotes (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================
        Tool(
            name="add_footnote",
            description="Add a footnote to a specific paragraph in the document.",
            inputSchema={
                "type": "object",
                "required": ["paragraph_index", "footnote_text"],
                "properties": {
                    "paragraph_index": {"type": "integer", "description": "1-indexed paragraph number"},
                    "footnote_text": {"type": "string", "description": "Text content of the footnote"}
                }
            }
        ),
        Tool(
            name="read_footnotes",
            description="Read all footnotes from the current document.",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="add_endnote",
            description="Add an endnote to a specific paragraph in the document.",
            inputSchema={
                "type": "object",
                "required": ["paragraph_index", "endnote_text"],
                "properties": {
                    "paragraph_index": {"type": "integer", "description": "1-indexed paragraph number"},
                    "endnote_text": {"type": "string", "description": "Text content of the endnote"}
                }
            }
        ),

        # ====================================================================
        # Advanced table formatting (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================
        Tool(
            name="format_table",
            description="Format a table with borders, header row styling, and cell shading.",
            inputSchema={
                "type": "object",
                "required": ["table_index"],
                "properties": {
                    "table_index": {"type": "integer", "description": "1-indexed table number"},
                    "has_header_row": {"type": "boolean", "description": "Bold the first row as header"},
                    "border_style": {"type": "string", "enum": ["none", "single", "double", "thick"]},
                    "shading": {"type": "array", "description": "2D list of hex colors (by row and column)", "items": {"type": "array", "items": {"type": "string"}}}
                }
            }
        ),
        Tool(
            name="set_table_cell_shading",
            description="Apply background shading to a specific table cell.",
            inputSchema={
                "type": "object",
                "required": ["table_index", "row_index", "col_index", "fill_color"],
                "properties": {
                    "table_index": {"type": "integer", "description": "1-indexed table number"},
                    "row_index": {"type": "integer", "description": "1-indexed row"},
                    "col_index": {"type": "integer", "description": "1-indexed column"},
                    "fill_color": {"type": "string", "description": "Hex color (e.g., FF0000) or color name"},
                    "pattern": {"type": "string", "description": "Shading pattern (clear, solid, pct10, etc.)", "default": "clear"}
                }
            }
        ),
        Tool(
            name="apply_table_alternating_rows",
            description="Apply alternating row colors to a table for better readability.",
            inputSchema={
                "type": "object",
                "required": ["table_index"],
                "properties": {
                    "table_index": {"type": "integer", "description": "1-indexed table number"},
                    "color1": {"type": "string", "description": "Hex color for odd rows (default FFFFFF)"},
                    "color2": {"type": "string", "description": "Hex color for even rows (default F2F2F2)"}
                }
            }
        ),
        Tool(
            name="highlight_table_header",
            description="Apply highlighting to the header row of a table.",
            inputSchema={
                "type": "object",
                "required": ["table_index"],
                "properties": {
                    "table_index": {"type": "integer", "description": "1-indexed table number"},
                    "header_color": {"type": "string", "description": "Background hex color (default 4472C4)"},
                    "text_color": {"type": "string", "description": "Text hex color (default FFFFFF)"}
                }
            }
        ),
        Tool(
            name="set_table_cell_alignment",
            description="Set text alignment for a specific table cell.",
            inputSchema={
                "type": "object",
                "required": ["table_index", "row_index", "col_index"],
                "properties": {
                    "table_index": {"type": "integer", "description": "1-indexed table number"},
                    "row_index": {"type": "integer", "description": "1-indexed row"},
                    "col_index": {"type": "integer", "description": "1-indexed column"},
                    "horizontal": {"type": "string", "enum": ["left", "center", "right", "justify"], "default": "left"},
                    "vertical": {"type": "string", "enum": ["top", "center", "bottom"], "default": "top"}
                }
            }
        ),
        Tool(
            name="set_table_column_width",
            description="Set the width of a specific table column.",
            inputSchema={
                "type": "object",
                "required": ["table_index", "col_index", "width"],
                "properties": {
                    "table_index": {"type": "integer", "description": "1-indexed table number"},
                    "col_index": {"type": "integer", "description": "1-indexed column"},
                    "width": {"type": "number", "description": "Width value"},
                    "width_type": {"type": "string", "enum": ["points", "inches", "cm", "percent", "auto"], "default": "points"}
                }
            }
        ),
        Tool(
            name="set_table_column_widths",
            description="Set widths for multiple table columns at once.",
            inputSchema={
                "type": "object",
                "required": ["table_index", "widths"],
                "properties": {
                    "table_index": {"type": "integer", "description": "1-indexed table number"},
                    "widths": {"type": "array", "items": {"type": "number"}, "description": "List of width values"},
                    "width_type": {"type": "string", "enum": ["points", "inches", "cm", "percent", "auto"], "default": "points"}
                }
            }
        ),
        Tool(
            name="set_table_cell_padding",
            description="Set cell padding/margins for a specific table cell.",
            inputSchema={
                "type": "object",
                "required": ["table_index", "row_index", "col_index"],
                "properties": {
                    "table_index": {"type": "integer", "description": "1-indexed table number"},
                    "row_index": {"type": "integer", "description": "1-indexed row"},
                    "col_index": {"type": "integer", "description": "1-indexed column"},
                    "top": {"type": "number", "description": "Top padding in points"},
                    "bottom": {"type": "number", "description": "Bottom padding in points"},
                    "left": {"type": "number", "description": "Left padding in points"},
                    "right": {"type": "number", "description": "Right padding in points"}
                }
            }
        ),

        # ====================================================================
        # Standalone list tools (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================
        Tool(
            name="insert_bulleted_list",
            description="Insert a bulleted list at the end of the document.",
            inputSchema={
                "type": "object",
                "required": ["items"],
                "properties": {
                    "items": {"type": "array", "items": {"type": "string"}, "description": "List item texts"},
                    "style": {"type": "string", "description": "List style name (default: List Bullet)"}
                }
            }
        ),
        Tool(
            name="insert_numbered_list",
            description="Insert a numbered list at the end of the document.",
            inputSchema={
                "type": "object",
                "required": ["items"],
                "properties": {
                    "items": {"type": "array", "items": {"type": "string"}, "description": "List item texts"},
                    "style": {"type": "string", "description": "List style name (default: List Number)"}
                }
            }
        ),

        # ====================================================================
        # Text highlighting (adapted from knorq-ai/docx-mcp-server)
        # ====================================================================
        Tool(
            name="highlight_text",
            description="Highlight all occurrences of specific text in the document.",
            inputSchema={
                "type": "object",
                "required": ["search_text"],
                "properties": {
                    "search_text": {"type": "string", "description": "Text to search for and highlight"},
                    "color": {"type": "string", "enum": ["yellow", "red", "green", "blue", "cyan", "magenta", "dark_yellow", "turquoise", "pink"], "default": "yellow"},
                    "match_case": {"type": "boolean", "description": "Case-sensitive match (default false)"}
                }
            }
        ),

        # ====================================================================
        # Document protection (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================
        Tool(
            name="protect_document",
            description="Encrypt the current document with a password. Requires msoffcrypto-tool.",
            inputSchema={
                "type": "object",
                "required": ["password"],
                "properties": {
                    "password": {"type": "string", "description": "Password to protect the document with"}
                }
            }
        ),
        Tool(
            name="unprotect_document",
            description="Decrypt a password-protected document. Requires msoffcrypto-tool.",
            inputSchema={
                "type": "object",
                "required": ["password"],
                "properties": {
                    "password": {"type": "string", "description": "Password to decrypt the document"}
                }
            }
        ),

        # ====================================================================
        # PDF export & Document merge (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================
        Tool(
            name="convert_to_pdf",
            description="Convert the current document to PDF. Uses docx2pdf on Windows, LibreOffice on Linux/macOS.",
            inputSchema={
                "type": "object",
                "properties": {
                    "output_path": {"type": "string", "description": "Output PDF path (defaults to same name as document)"}
                }
            }
        ),
        Tool(
            name="merge_documents",
            description="Merge multiple .docx files into a single document.",
            inputSchema={
                "type": "object",
                "required": ["source_paths", "output_path"],
                "properties": {
                    "source_paths": {"type": "array", "items": {"type": "string"}, "description": "List of source .docx file paths (in merge order)"},
                    "output_path": {"type": "string", "description": "Output path for merged document"}
                }
            }
        ),
        # ====================================================================
        # Stable anchors, images, comment threading (adapted from knorq-ai)
        # ====================================================================
        Tool(
            name="ensure_anchors",
            description="Assign stable w14:paraId anchors to every top-level paragraph. Anchors survive index shifts from insert/delete. Idempotent.",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="list_images",
            description="List all images embedded in the current document with dimensions and relationship info.",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="reply_to_comment",
            description="Reply to an existing comment, creating a threaded conversation.",
            inputSchema={
                "type": "object",
                "required": ["comment_id", "reply_text"],
                "properties": {
                    "comment_id": {"type": "integer", "description": "ID of the parent comment to reply to"},
                    "reply_text": {"type": "string", "description": "The reply content"},
                    "author": {"type": "string", "description": "Reply author name", "default": "Claude"}
                }
            }
        ),
        Tool(
            name="delete_comment",
            description="Delete a comment by its ID. Also removes range markers from the document.",
            inputSchema={
                "type": "object",
                "required": ["comment_id"],
                "properties": {
                    "comment_id": {"type": "integer", "description": "Comment ID to delete"}
                }
            }
        ),
        Tool(
            name="get_page_layout",
            description="Get page size, margins, orientation, and detected preset of the current document.",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="format_text",
            description="Apply character formatting (bold, italic, underline, strikethrough, highlight, font, size, color) to all runs matching search text.",
            inputSchema={
                "type": "object",
                "required": ["search"],
                "properties": {
                    "search": {"type": "string", "description": "Text to find and format"},
                    "bold": {"type": "boolean", "description": "Set bold"},
                    "italic": {"type": "boolean", "description": "Set italic"},
                    "underline": {"type": "boolean", "description": "Set underline"},
                    "strikethrough": {"type": "boolean", "description": "Set strikethrough"},
                    "highlight_color": {"type": "string", "description": "Highlight color: yellow, green, cyan, magenta, blue, red, etc."},
                    "font_name": {"type": "string", "description": "Font family name"},
                    "font_size": {"type": "number", "description": "Font size in points"},
                    "font_color": {"type": "string", "description": "Font color as hex (e.g. 'FF0000' for red)"},
                    "match_case": {"type": "boolean", "description": "Case-sensitive matching", "default": False}
                }
            }
        ),
        Tool(
            name="set_headings",
            description="Convert multiple paragraphs to headings in one operation.",
            inputSchema={
                "type": "object",
                "required": ["headings"],
                "properties": {
                    "headings": {
                        "type": "array",
                        "description": "Array of heading assignments",
                        "items": {
                            "type": "object",
                            "properties": {
                                "paragraph_index": {"type": "integer", "description": "1-based paragraph index"},
                                "level": {"type": "integer", "description": "Heading level (1-9)", "default": 1}
                            },
                            "required": ["paragraph_index", "level"]
                        }
                    }
                }
            }
        ),
        Tool(
            name="get_paragraph_format",
            description="Introspect a paragraph's formatting: style, alignment, spacing, indentation, and run-level formatting.",
            inputSchema={
                "type": "object",
                "required": ["paragraph_index"],
                "properties": {
                    "paragraph_index": {"type": "integer", "description": "1-based paragraph index"}
                }
            }
        ),

        # ---- Bibliography / Sources ----
        # WORKFLOW: To add citations to a document, follow these steps IN ORDER:
        #   1. add_bibliography_source(s)  — add sources to the master list (Sources.xml)
        #   2. open_document              — open the target .docx file
        #   3. copy_sources_to_current_list — copy master sources into the document's Current List
        #      (CITATION fields will NOT resolve without this step!)
        #   4. insert_citation            — replace typed citations with Word CITATION fields
        #   5. insert_bibliography         — add auto-updating BIBLIOGRAPHY field at end
        #   6. User presses Ctrl+A then F9 in Word to render all fields
        Tool(
            name="add_bibliography_source",
            description=(
                "STEP 1 of bibliography workflow: Add a single citation source to Word's master bibliography "
                "(Sources.xml at %APPDATA%/Microsoft/Bibliography/Sources.xml on Windows, "
                "~/Library/Application Support/Microsoft/Bibliography/Sources.xml on macOS). "
                "These sources appear in Word under References > Manage Sources. "
                "After adding sources, use copy_sources_to_current_list to make them available in the document. "
                "Workflow: 1.add_bibliography_source(s) → 2.open_document → 3.copy_sources_to_current_list → 4.insert_citation → 5.insert_bibliography"
            ),
            inputSchema={
                "type": "object",
                "required": ["source_type", "title", "year"],
                "properties": {
                    "source_type": {
                        "type": "string",
                        "enum": ["JournalArticle", "Book", "Report", "ConferenceProceedings",
                                 "BookSection", "MagazineArticle", "NewspaperArticle",
                                 "WebSite", "DocumentFromInternetSite", "ElectronicSource",
                                 "Film", "Interview", "Patent", "Thesis"],
                        "description": "The type of the source"
                    },
                    "tag": {"type": "string", "description": "Unique tag/identifier for the source (auto-generated from author+year if omitted)"},
                    "title": {"type": "string", "description": "Title of the work"},
                    "year": {"type": "string", "description": "Publication year"},
                    "authors": {
                        "type": "array",
                        "description": "List of authors. Each author is an object with 'last' and 'first' keys, or {'corporate': 'OrgName'} for corporate authors.",
                        "items": {
                            "type": "object",
                            "properties": {
                                "last": {"type": "string", "description": "Last name"},
                                "first": {"type": "string", "description": "First name / initials"},
                                "corporate": {"type": "string", "description": "Corporate/organization author name (use instead of last/first)"                            }
                            }
                        }
                    },
                    "journal_name": {"type": "string", "description": "Journal name (for JournalArticle, MagazineArticle, NewspaperArticle)"},
                    "book_title": {"type": "string", "description": "Book title (for BookSection)"},
                    "volume": {"type": "string", "description": "Volume number"},
                    "issue": {"type": "string", "description": "Issue number"},
                    "pages": {"type": "string", "description": "Page range, e.g. '1349-1370'"},
                    "publisher": {"type": "string", "description": "Publisher name (for Book, Report)"},
                    "city": {"type": "string", "description": "Publication city (for Book, Report)"},
                    "doi": {"type": "string", "description": "DOI (e.g. '10.1234/abc')"},
                    "url": {"type": "string", "description": "URL"},
                    "abstract": {"type": "string", "description": "Abstract"},
                    "edition": {"type": "string", "description": "Edition (for Book)"},
                    "isbn": {"type": "string", "description": "ISBN"},
                    "issn": {"type": "string", "description": "ISSN"},
                    "language": {"type": "string", "description": "Language"},
                    "comments": {"type": "string", "description": "Comments / notes"},
                    "conference_name": {"type": "string", "description": "Conference name (for ConferenceProceedings)"},
                    "medium": {"type": "string", "description": "Medium (for ElectronicSource, Film)"},
                    "number": {"type": "string", "description": "Patent number, report number, etc."},
                    "country": {"type": "string", "description": "Country (for Patent)"},
                    "institution": {"type": "string", "description": "Institution (for Thesis)"},
                    "department": {"type": "string", "description": "Department (for Thesis)"},
                    "thesis_type": {"type": "string", "description": "Thesis type: Master's or Ph.D. (for Thesis)"},
                    "access_date": {"type": "string", "description": "Date accessed (for web sources)"},
                    "production_company": {"type": "string", "description": "Production company (for Film)"},
                    "guid": {"type": "string", "description": "GUID for the source (auto-generated if omitted)"}
                }
            }
        ),
        Tool(
            name="add_bibliography_sources_batch",
            description=(
                "STEP 1 of bibliography workflow: Add multiple citation sources to Word's master bibliography in one call. "
                "Each source follows the same schema as add_bibliography_source. "
                "Use this for bulk-importing reference lists. "
                "After adding, use copy_sources_to_current_list to copy them into the document. "
                "Workflow: 1.add_bibliography_source(s) → 2.open_document → 3.copy_sources_to_current_list → 4.insert_citation → 5.insert_bibliography"
            ),
            inputSchema={
                "type": "object",
                "required": ["sources"],
                "properties": {
                    "sources": {
                        "type": "array",
                        "description": "Array of source objects (same schema as add_bibliography_source)",
                        "items": {
                            "type": "object",
                            "required": ["source_type", "title", "year"],
                            "properties": {
                                "source_type": {
                                    "type": "string",
                                    "enum": ["JournalArticle", "Book", "Report", "ConferenceProceedings",
                                             "BookSection", "MagazineArticle", "NewspaperArticle",
                                             "WebSite", "DocumentFromInternetSite", "ElectronicSource",
                                             "Film", "Interview", "Patent", "Thesis"],
                                    "description": "The type of the source"
                                },
                                "tag": {"type": "string", "description": "Unique tag (auto-generated if omitted)"},
                                "title": {"type": "string", "description": "Title of the work"},
                                "year": {"type": "string", "description": "Publication year"},
                                "authors": {
                                    "type": "array",
                                    "description": "List of authors. Each: {'last':'...','first':'...'} or {'corporate':'...'}",
                                    "items": {
                                        "type": "object",
                                        "properties": {
                                            "last": {"type": "string"},
                                            "first": {"type": "string"},
                                            "corporate": {"type": "string"}
                                        }
                                    }
                                },
                                "journal_name": {"type": "string"},
                                "book_title": {"type": "string"},
                                "volume": {"type": "string"},
                                "issue": {"type": "string"},
                                "pages": {"type": "string"},
                                "publisher": {"type": "string"},
                                "city": {"type": "string"},
                                "doi": {"type": "string"},
                                "url": {"type": "string"},
                                "abstract": {"type": "string"},
                                "edition": {"type": "string"},
                                "isbn": {"type": "string"},
                                "issn": {"type": "string"},
                                "language": {"type": "string"},
                                "comments": {"type": "string"},
                                "conference_name": {"type": "string"},
                                "medium": {"type": "string"},
                                "number": {"type": "string"},
                                "country": {"type": "string"},
                                "institution": {"type": "string"},
                                "department": {"type": "string"},
                                "thesis_type": {"type": "string"},
                                "access_date": {"type": "string"},
                                "production_company": {"type": "string"},
                                "guid": {"type": "string"}
                            }
                        }
                    }
                }
            }
        ),
        Tool(
            name="list_bibliography_sources",
            description=(
                "List all citation sources currently in Word's master bibliography (Sources.xml). "
                "Returns each source's tag, type, title, and authors. "
                "Use this to find source tags needed for insert_citation."
            ),
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="clear_bibliography_sources",
            description=(
                "Remove all citation sources from Word's master bibliography (Sources.xml). "
                "WARNING: This deletes all sources from the MASTER list — use with caution. "
                "This does NOT remove sources from individual documents' Current Lists."
            ),
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="insert_citation",
            description=(
                "STEP 4 of bibliography workflow: Replace typed citation text in the document with a proper Word CITATION field "
                "linked to source(s) in the bibliography. The field will render as a proper "
                "in-text citation (e.g. '(Author, Year)') when the user presses Ctrl+A, F9 in Word. "
                "PREREQUISITE: Sources must be in the document's Current List — run copy_sources_to_current_list FIRST, "
                "otherwise Word will show 'Invalid source specified'. "
                "Use list_bibliography_sources to find the correct source tags. "
                "Workflow: 1.add_bibliography_source(s) → 2.open_document → 3.copy_sources_to_current_list → 4.insert_citation → 5.insert_bibliography"
            ),
            inputSchema={
                "type": "object",
                "required": ["source_tags"],
                "properties": {
                    "source_tags": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Array of source tags (matching Tags in Sources.xml) to cite"
                    },
                    "match_text": {
                        "type": "string",
                        "description": "The typed citation text to find and replace (e.g. '(Perera & Wijesinghe, 2025)'). If omitted, citation is appended to the paragraph."
                    },
                    "paragraph_index": {
                        "type": "integer",
                        "description": "1-based paragraph index to search in. If omitted, searches all paragraphs."
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "1-based table index. If provided with row/col, replaces the cell content with the citation."
                    },
                    "row": {
                        "type": "integer",
                        "description": "1-based row index in the table (requires table_index)"
                    },
                    "col": {
                        "type": "integer",
                        "description": "1-based column index in the table (requires table_index)"
                    }
                }
            }
        ),
        Tool(
            name="insert_citations_batch",
            description=(
                "STEP 4 of bibliography workflow (batch): Replace multiple typed citation texts in one call. "
                "Each item in the citations array can target either a paragraph (via match_text) or a table cell (via table_index/row/col). "
                "This is the batch version of insert_citation — use it to replace all typed citations at once. "
                "PREREQUISITE: Run copy_sources_to_current_list FIRST. "
                "Workflow: 1.add_bibliography_source(s) → 2.open_document → 3.copy_sources_to_current_list → 4.insert_citations_batch → 5.insert_bibliography"
            ),
            inputSchema={
                "type": "object",
                "required": ["citations"],
                "properties": {
                    "citations": {
                        "type": "array",
                        "description": "Array of citation replacement objects",
                        "items": {
                            "type": "object",
                            "required": ["source_tags"],
                            "properties": {
                                "source_tags": {
                                    "type": "array",
                                    "items": {"type": "string"},
                                    "description": "Source tags to cite"
                                },
                                "match_text": {
                                    "type": "string",
                                    "description": "Typed citation text to find and replace (for paragraph mode)"
                                },
                                "paragraph_index": {
                                    "type": "integer",
                                    "description": "1-based paragraph index (optional, searches all if omitted)"
                                },
                                "table_index": {
                                    "type": "integer",
                                    "description": "1-based table index (for table cell mode)"
                                },
                                "row": {
                                    "type": "integer",
                                    "description": "1-based row index (requires table_index)"
                                },
                                "col": {
                                    "type": "integer",
                                    "description": "1-based column index (requires table_index)"
                                }
                            }
                        }
                    }
                }
            }
        ),
        Tool(
            name="copy_sources_to_current_list",
            description=(
                "STEP 3 of bibliography workflow: Copy all sources from Word's master bibliography (Sources.xml) into the "
                "currently open document's Current List (customXml/item1.xml inside the .docx). "
                "CRITICAL: Word's CITATION fields only resolve against the Current List, NOT the Master List. "
                "Without this step, insert_citation will produce 'Invalid source specified' errors in Word. "
                "Run this AFTER add_bibliography_sources_batch and AFTER open_document, but BEFORE insert_citation. "
                "Workflow: 1.add_bibliography_source(s) → 2.open_document → 3.copy_sources_to_current_list → 4.insert_citation → 5.insert_bibliography"
            ),
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="insert_bibliography",
            description=(
                "STEP 5 of bibliography workflow: Insert an auto-updating BIBLIOGRAPHY field at the end of the document. "
                "This generates the reference list from all cited sources. "
                "The user must press Ctrl+A then F9 in Word to populate it. "
                "Run this AFTER insert_citation so all cited sources are in the document. "
                "Workflow: 1.add_bibliography_source(s) → 2.open_document → 3.copy_sources_to_current_list → 4.insert_citation → 5.insert_bibliography"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "heading_text": {
                        "type": "string",
                        "description": "Heading text for the bibliography section (default: 'References')"
                    },
                    "heading_style": {
                        "type": "string",
                        "description": "Paragraph style for the heading (default: 'Heading 1')"
                    }
                }
            }
        ),
    ]


# ---------------------------------------------------------------------------
# Tool implementations
# ---------------------------------------------------------------------------

@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    global _current_doc_path, _doc

    def ok(msg):
        return [TextContent(type="text", text=str(msg))]

    # Shared alignment map
    ALIGN_MAP = {
        "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
        "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
        "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
        "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }

    try:

        # ====================================================================
        # Document lifecycle
        # ====================================================================

        if name == "open_document":
            path = arguments["path"]
            if not os.path.exists(path):
                return ok(f"ERROR: File not found: {path}")
            _current_doc_path = path
            _doc = Document(path)
            return ok(
                f"Opened: {path}\n"
                f"Paragraphs: {len(_doc.paragraphs)} | "
                f"Tables: {len(_doc.tables)} | "
                f"Sections: {len(_doc.sections)}"
            )

        elif name == "create_new_document":
            path = arguments["path"]
            template = arguments.get("template")
            doc = Document(template) if template else Document()
            # Apply default font/size if specified
            if arguments.get("default_font") or arguments.get("default_font_size"):
                style = doc.styles["Normal"]
                if arguments.get("default_font"):
                    style.font.name = arguments["default_font"]
                if arguments.get("default_font_size"):
                    style.font.size = Pt(arguments["default_font_size"])
            doc.save(path)
            _current_doc_path = path
            _doc = Document(path)
            return ok(f"Created: {path}")

        elif name == "save_document":
            backup = arguments.get("backup", True)
            _save(backup=backup)
            return ok(f"Saved: {_current_doc_path}")

        elif name == "save_as":
            path = arguments["path"]
            _require_doc().save(path)
            _current_doc_path = path
            _reload_in_word(path)
            return ok(f"Saved as: {path}")

        elif name == "close_document":
            _doc = None
            _current_doc_path = None
            return ok("Document closed.")

        elif name == "duplicate_document":
            doc = _require_doc()
            new_path = arguments["new_path"]
            shutil.copy2(_current_doc_path, new_path)
            return ok(f"Duplicated to: {new_path}")

        # ====================================================================
        # Read & inspect
        # ====================================================================

        elif name == "get_document_info":
            doc = _require_doc()
            cp = doc.core_properties
            words = sum(len(p.text.split()) for p in doc.paragraphs)
            chars = sum(len(p.text) for p in doc.paragraphs)
            result = {
                "path": _current_doc_path,
                "title": cp.title,
                "author": cp.author,
                "last_modified_by": cp.last_modified_by,
                "created": str(cp.created),
                "modified": str(cp.modified),
                "subject": cp.subject,
                "keywords": cp.keywords,
                "description": cp.description,
                "category": cp.category,
                "paragraphs": len(doc.paragraphs),
                "tables": len(doc.tables),
                "sections": len(doc.sections),
                "word_count": words,
                "char_count": chars,
            }
            return ok(json.dumps(result, indent=2, default=str))

        elif name == "read_document":
            doc = _require_doc()
            include_xml = arguments.get("include_xml", False)
            include_tables = arguments.get("include_tables", True)
            elements = []
            for i, p in enumerate(doc.paragraphs):
                el = {
                    "index": i,
                    "type": "paragraph",
                    "style": _para_style(p),
                    "text": p.text,
                    "alignment": str(p.alignment),
                    "runs": [
                        {
                            "text": r.text,
                            "bold": r.bold,
                            "italic": r.italic,
                            "underline": r.underline,
                            "font_size": r.font.size.pt if r.font.size else None,
                            "font_name": r.font.name,
                            "color": str(r.font.color.rgb) if r.font.color and r.font.color.type else None,
                        }
                        for r in p.runs
                    ],
                }
                if include_xml:
                    el["xml"] = _xml_to_str(p._p)
                elements.append(el)
            tables = []
            if include_tables:
                for ti, tbl in enumerate(doc.tables):
                    rows_data = []
                    for row in tbl.rows:
                        rows_data.append([cell.text for cell in row.cells])
                    tables.append({
                        "table_index": ti,
                        "style": tbl.style.name if tbl.style else None,
                        "rows": rows_data
                    })
            return ok(json.dumps({"paragraphs": elements, "tables": tables}, indent=2, default=str))

        elif name == "get_outline":
            doc = _require_doc()
            outline = []
            for i, p in enumerate(doc.paragraphs):
                lvl = _heading_level(p)
                if lvl is not None:
                    outline.append({
                        "index": i, "level": lvl, "text": p.text
                    })
            lines = [
                f"[{o['index']}] {'  ' * (o['level'] - 1)}H{o['level']}: {o['text']}"
                for o in outline
            ]
            return ok("\n".join(lines) if lines else "No headings found.")

        elif name == "get_section":
            doc = _require_doc()
            heading = arguments["heading"]
            idxs = _find_paras_under_heading(doc, heading)
            if not idxs:
                return ok(f"Heading '{heading}' not found or section is empty.")
            paras = doc.paragraphs
            lines = [f"[{i}] ({_para_style(paras[i])}) {paras[i].text}" for i in idxs]
            return ok("\n".join(lines))

        elif name == "find_text":
            doc = _require_doc()
            query = arguments["query"]
            case_sensitive = arguments.get("case_sensitive", False)
            include_tables = arguments.get("include_tables", True)
            results = []
            for i, p in enumerate(doc.paragraphs):
                haystack = p.text if case_sensitive else p.text.lower()
                needle = query if case_sensitive else query.lower()
                if needle in haystack:
                    results.append({"source": "paragraph", "index": i,
                                    "style": _para_style(p), "text": p.text})
            if include_tables:
                for ti, tbl in enumerate(doc.tables):
                    for ri, row in enumerate(tbl.rows):
                        for ci, cell in enumerate(row.cells):
                            haystack = cell.text if case_sensitive else cell.text.lower()
                            needle = query if case_sensitive else query.lower()
                            if needle in haystack:
                                results.append({
                                    "source": "table",
                                    "table_index": ti,
                                    "row": ri, "col": ci,
                                    "text": cell.text
                                })
            if not results:
                return ok(f"No occurrences of '{query}' found.")
            return ok(json.dumps(results, indent=2))

        elif name == "read_table":
            doc = _require_doc()
            ti = arguments.get("table_index")
            tables = doc.tables
            if ti is not None:
                selected = [tables[ti]]
                base_idx = ti
            else:
                selected = tables
                base_idx = 0
            result = []
            for idx, tbl in enumerate(selected):
                rows_data = [[cell.text for cell in row.cells] for row in tbl.rows]
                result.append({
                    "index": idx + base_idx,
                    "style": tbl.style.name if tbl.style else None,
                    "rows": len(tbl.rows),
                    "cols": len(tbl.columns),
                    "data": rows_data
                })
            return ok(json.dumps(result, indent=2))

        elif name == "list_styles":
            doc = _require_doc()
            stype = arguments.get("style_type", "paragraph")
            type_map = {
                "paragraph": WD_STYLE_TYPE.PARAGRAPH,
                "character": WD_STYLE_TYPE.CHARACTER,
                "table": WD_STYLE_TYPE.TABLE,
            }
            if stype == "all":
                styles = [{"name": s.name, "type": str(s.type), "id": s.style_id}
                          for s in doc.styles]
            else:
                styles = [{"name": s.name, "id": s.style_id}
                          for s in doc.styles if s.type == type_map.get(stype)]
            return ok(json.dumps(styles, indent=2))

        elif name == "get_document_xml":
            doc = _require_doc()
            if arguments.get("full_document"):
                return ok(_xml_to_str(doc.element.body))
            idx = arguments.get("paragraph_index")
            if idx is not None:
                return ok(_xml_to_str(doc.paragraphs[idx]._p))
            return ok("Specify paragraph_index or full_document=true")

        elif name == "read_headers_footers":
            doc = _require_doc()
            result = []
            for si, section in enumerate(doc.sections):
                entry = {"section": si}
                for which in ["header", "footer", "first_page_header", "first_page_footer",
                              "even_page_header", "even_page_footer"]:
                    try:
                        hf = getattr(section, which)
                        if hf and not hf.is_linked_to_previous:
                            text = "\n".join(p.text for p in hf.paragraphs)
                            entry[which] = text
                    except Exception:
                        pass
                result.append(entry)
            return ok(json.dumps(result, indent=2))

        # ====================================================================
        # Text editing
        # ====================================================================

        elif name == "replace_text":
            doc = _require_doc()
            find = arguments["find"]
            replace_with = arguments["replace"]
            case_sensitive = arguments.get("case_sensitive", False)
            whole_word = arguments.get("whole_word", False)
            include_tables = arguments.get("include_tables", True)
            count = 0

            def do_replace(runs):
                nonlocal count
                for run in runs:
                    orig = run.text
                    text_cmp = orig if case_sensitive else orig.lower()
                    needle = find if case_sensitive else find.lower()
                    if whole_word:
                        pattern = rf"\b{re.escape(needle)}\b"
                    else:
                        pattern = re.escape(needle)
                    flags = 0 if case_sensitive else re.IGNORECASE
                    if re.search(pattern, text_cmp):
                        run.text = re.sub(
                            rf"\b{re.escape(find)}\b" if whole_word else re.escape(find),
                            replace_with, orig, flags=flags
                        )
                        count += 1

            for para in doc.paragraphs:
                do_replace(para.runs)
            if include_tables:
                for tbl in doc.tables:
                    for row in tbl.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                do_replace(para.runs)
            _save()
            return ok(f"Replaced {count} occurrence(s) of '{find}' -> '{replace_with}'")

        elif name == "replace_text_batch":
            doc = _require_doc()
            replacements = arguments.get("replacements", [])
            if not replacements:
                return ok("No replacements provided.")
            total_count = 0
            summary = []
            for item in replacements:
                find = item["find"]
                replace_with = item["replace"]
                case_sensitive = item.get("case_sensitive", False)
                whole_word = item.get("whole_word", False)
                include_tables = item.get("include_tables", True)
                count = 0

                def do_replace(runs):
                    nonlocal count
                    for run in runs:
                        orig = run.text
                        text_cmp = orig if case_sensitive else orig.lower()
                        needle = find if case_sensitive else find.lower()
                        if whole_word:
                            pattern = rf"\b{re.escape(needle)}\b"
                        else:
                            pattern = re.escape(needle)
                        flags = 0 if case_sensitive else re.IGNORECASE
                        if re.search(pattern, text_cmp):
                            run.text = re.sub(
                                rf"\b{re.escape(find)}\b" if whole_word else re.escape(find),
                                replace_with, orig, flags=flags
                            )
                            count += 1

                for para in doc.paragraphs:
                    do_replace(para.runs)
                if include_tables:
                    for tbl in doc.tables:
                        for row in tbl.rows:
                            for cell in row.cells:
                                for para in cell.paragraphs:
                                    do_replace(para.runs)
                total_count += count
                summary.append(f"'{find}' -> '{replace_with}': {count} occurrence(s)")
            _save()
            return ok(f"Batch complete: {total_count} total replacement(s).\n" + "\n".join(summary))

        elif name == "replace_paragraph":
            doc = _require_doc()
            new_text = arguments["new_text"]
            target = None
            if "paragraph_index" in arguments:
                target = doc.paragraphs[arguments["paragraph_index"]]
            elif "match_text" in arguments:
                for p in doc.paragraphs:
                    if arguments["match_text"].lower() in p.text.lower():
                        target = p
                        break
            if target is None:
                return ok("Paragraph not found.")
            for run in target.runs:
                run.text = ""
            if target.runs:
                target.runs[0].text = new_text
            else:
                target.add_run(new_text)
            _save()
            return ok("Paragraph replaced.")

        elif name == "replace_section":
            doc = _require_doc()
            heading = arguments["heading"]
            new_paragraphs = arguments["new_paragraphs"]
            idxs = _find_paras_under_heading(doc, heading)
            if not idxs:
                return ok(f"Heading '{heading}' not found.")
            paras = doc.paragraphs
            for i in reversed(idxs):
                _delete_paragraph(paras[i])
            # Re-find heading after deletions
            heading_para = _find_heading_para(doc, heading)
            if heading_para is None:
                return ok("Heading paragraph no longer found after deletion.")
            ref = heading_para
            for np_data in new_paragraphs:
                style_name = np_data.get("style", "Normal")
                style_id = _get_style_id(doc, style_name) or "Normal"
                _insert_paragraph_after(ref, np_data["text"], style_id)
                # Advance ref to the newly inserted paragraph
                idx_ref = list(doc.paragraphs).index(ref)
                if idx_ref + 1 < len(doc.paragraphs):
                    ref = doc.paragraphs[idx_ref + 1]
            _save()
            return ok(f"Section '{heading}' replaced with {len(new_paragraphs)} paragraph(s).")

        elif name == "insert_paragraph":
            doc = _require_doc()
            text = arguments["text"]
            style = arguments.get("style", "Normal")
            style_id = _get_style_id(doc, style) or "Normal"
            paras = doc.paragraphs

            ref_para = None
            if "after_text" in arguments:
                for p in paras:
                    if arguments["after_text"].lower() in p.text.lower():
                        ref_para = p
                        break
            elif "after_heading" in arguments:
                ref_para = _find_heading_para(doc, arguments["after_heading"])
            elif "at_index" in arguments:
                idx = arguments["at_index"]
                if 0 < idx < len(paras):
                    ref_para = paras[idx - 1]

            if ref_para is not None:
                new_p = _insert_paragraph_after(ref_para, text, style_id)
                # Find the newly created paragraph object
                new_para = None
                for p in doc.paragraphs:
                    if p._p is new_p:
                        new_para = p
                        break
            else:
                # Append at end
                new_para = doc.add_paragraph(text)
                try:
                    new_para.style = doc.styles[style]
                except Exception:
                    pass

            # Apply optional run-level formatting if paragraph found
            if new_para:
                align_str = arguments.get("alignment")
                if align_str:
                    new_para.alignment = ALIGN_MAP.get(align_str.upper(), WD_ALIGN_PARAGRAPH.LEFT)
                for run in new_para.runs:
                    _set_run_formatting(
                        run,
                        bold=arguments.get("bold"),
                        italic=arguments.get("italic"),
                        font_size=arguments.get("font_size"),
                        font_name=arguments.get("font_name"),
                        color_hex=arguments.get("color_hex"),
                    )
            _save()
            return ok("Paragraph inserted.")

        elif name == "delete_paragraph":
            doc = _require_doc()
            paras = doc.paragraphs
            if "paragraph_index" in arguments:
                _delete_paragraph(paras[arguments["paragraph_index"]])
                _save()
                return ok("Paragraph deleted.")
            if "match_text" in arguments:
                delete_all = arguments.get("delete_all_matching", False)
                to_delete = []
                for p in paras:
                    if arguments["match_text"].lower() in p.text.lower():
                        to_delete.append(p)
                        if not delete_all:
                            break
                for p in to_delete:
                    _delete_paragraph(p)
                _save()
                return ok(f"Deleted {len(to_delete)} paragraph(s).")
            return ok("Specify match_text or paragraph_index.")

        elif name == "delete_section":
            doc = _require_doc()
            heading = arguments["heading"]
            heading_para = _find_heading_para(doc, heading)
            if heading_para is None:
                return ok(f"Heading '{heading}' not found.")
            idxs = _find_paras_under_heading(doc, heading)
            for i in reversed(idxs):
                _delete_paragraph(doc.paragraphs[i])
            _delete_paragraph(heading_para)
            _save()
            return ok(f"Section '{heading}' deleted ({len(idxs) + 1} paragraphs removed).")

        elif name == "move_section":
            doc = _require_doc()
            section_heading = arguments["section_heading"]
            before_heading = arguments.get("before_heading")
            after_heading_arg = arguments.get("after_heading")

            heading_para = _find_heading_para(doc, section_heading)
            if heading_para is None:
                return ok(f"Section heading '{section_heading}' not found.")
            idxs = _find_paras_under_heading(doc, section_heading)
            all_paras = [heading_para] + [doc.paragraphs[i] for i in idxs]
            clones = [copy.deepcopy(p._p) for p in all_paras]
            for p in reversed(all_paras):
                _delete_paragraph(p)

            target_text = before_heading or after_heading_arg
            target_para = _find_heading_para(doc, target_text)
            if target_para is None:
                return ok(f"Target heading '{target_text}' not found.")

            if before_heading:
                ref = target_para._p
                for clone in reversed(clones):
                    ref.addprevious(clone)
            else:
                ref = target_para._p
                for clone in clones:
                    ref.addnext(clone)
                    ref = clone
            _save()
            return ok(f"Section '{section_heading}' moved.")

        # ====================================================================
        # Formatting
        # ====================================================================

        elif name == "format_paragraph":
            doc = _require_doc()
            target = None
            if "paragraph_index" in arguments:
                target = doc.paragraphs[arguments["paragraph_index"]]
            elif "match_text" in arguments:
                for p in doc.paragraphs:
                    if arguments["match_text"].lower() in p.text.lower():
                        target = p
                        break
            if target is None:
                return ok("Paragraph not found.")

            if "style" in arguments:
                try:
                    target.style = doc.styles[arguments["style"]]
                except KeyError:
                    pass
            if "alignment" in arguments:
                target.alignment = ALIGN_MAP.get(arguments["alignment"].upper(), WD_ALIGN_PARAGRAPH.LEFT)

            pf = target.paragraph_format
            if "space_before" in arguments:
                pf.space_before = Pt(arguments["space_before"])
            if "space_after" in arguments:
                pf.space_after = Pt(arguments["space_after"])
            if "left_indent" in arguments:
                pf.left_indent = Cm(arguments["left_indent"])
            if "right_indent" in arguments:
                pf.right_indent = Cm(arguments["right_indent"])
            if "first_line_indent" in arguments:
                pf.first_line_indent = Cm(arguments["first_line_indent"])
            if "line_spacing" in arguments:
                pf.line_spacing = Pt(arguments["line_spacing"])
            if "keep_together" in arguments:
                pf.keep_together = arguments["keep_together"]
            if "keep_with_next" in arguments:
                pf.keep_with_next = arguments["keep_with_next"]
            if "page_break_before" in arguments:
                pf.page_break_before = arguments["page_break_before"]

            _save()
            return ok("Paragraph formatted.")

        elif name == "format_text_run":
            doc = _require_doc()
            para_match = arguments.get("paragraph_match", "")
            para_idx = arguments.get("paragraph_index")
            run_match = arguments.get("run_text_match", "")

            targets = []
            if para_idx is not None:
                targets = [doc.paragraphs[para_idx]]
            elif para_match:
                for p in doc.paragraphs:
                    if para_match.lower() in p.text.lower():
                        targets.append(p)

            for p in targets:
                for run in p.runs:
                    if not run_match or run_match.lower() in run.text.lower():
                        _set_run_formatting(
                            run,
                            bold=arguments.get("bold"),
                            italic=arguments.get("italic"),
                            underline=arguments.get("underline"),
                            strike=arguments.get("strike"),
                            font_size=arguments.get("font_size"),
                            font_name=arguments.get("font_name"),
                            color_hex=arguments.get("color_hex"),
                            all_caps=arguments.get("all_caps"),
                            small_caps=arguments.get("small_caps"),
                        )
            _save()
            return ok(f"Run formatting applied to {len(targets)} paragraph(s).")

        # ====================================================================
        # Headings
        # ====================================================================

        elif name == "add_heading":
            doc = _require_doc()
            text = arguments["text"]
            level = arguments.get("level", 1)
            after_heading = arguments.get("after_heading")

            if after_heading:
                ref = _find_heading_para(doc, after_heading)
                if ref:
                    style_name = f"Heading {level}"
                    style_id = _get_style_id(doc, style_name) or f"Heading{level}"
                    _insert_paragraph_after(ref, text, style_id)
                    _save()
                    return ok(f"Heading '{text}' (H{level}) inserted after '{after_heading}'.")

            p = doc.add_heading(text, level=level)
            _save()
            return ok(f"Heading '{text}' (H{level}) added.")

        # ====================================================================
        # Tables
        # ====================================================================

        elif name == "insert_table":
            doc = _require_doc()
            rows = arguments["rows"]
            cols = arguments["cols"]
            data = arguments.get("data", [])
            style = arguments.get("style", "Table Grid")
            header_row = arguments.get("header_row", False)
            col_widths = arguments.get("col_widths", [])
            after_text = arguments.get("after_text")
            after_heading = arguments.get("after_heading")

            # Find anchor paragraph for positioning
            anchor_para = None
            if after_heading:
                anchor_para = _find_heading_para(doc, after_heading)
            elif after_text:
                for p in doc.paragraphs:
                    if after_text.lower() in p.text.lower():
                        anchor_para = p
                        break

            # python-docx always appends tables at the end of the body.
            # We add it at end first, then move the XML element to the correct position.
            tbl = doc.add_table(rows=rows, cols=cols)
            try:
                tbl.style = doc.styles[style]
            except Exception:
                pass

            for ri, row_data in enumerate(data):
                if ri >= rows:
                    break
                for ci, cell_val in enumerate(row_data):
                    if ci >= cols:
                        break
                    cell = tbl.cell(ri, ci)
                    if isinstance(cell_val, dict):
                        cell.text = ""
                        run = cell.paragraphs[0].add_run(str(cell_val.get("text", "")))
                        _set_run_formatting(
                            run,
                            bold=cell_val.get("bold"),
                            italic=cell_val.get("italic"),
                            color_hex=cell_val.get("color_hex"),
                            font_size=cell_val.get("font_size"),
                        )
                        if cell_val.get("bg_color"):
                            _set_cell_background(cell, cell_val["bg_color"])
                    else:
                        cell.text = str(cell_val)

                    if header_row and ri == 0:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.bold = True

            if col_widths:
                for ci, width_cm in enumerate(col_widths):
                    if ci < cols:
                        for row in tbl.rows:
                            row.cells[ci].width = Cm(width_cm)

            # Move table XML element immediately after the anchor paragraph
            if anchor_para is not None:
                tbl._tbl.getparent().remove(tbl._tbl)
                anchor_para._p.addnext(tbl._tbl)

            _save()
            return ok(f"Table {rows}x{cols} inserted.")

        elif name == "edit_table_cell":
            doc = _require_doc()
            ti = arguments["table_index"]
            row = arguments["row"]
            col = arguments["col"]
            text = arguments["text"]
            cell = doc.tables[ti].cell(row, col)
            cell.text = text
            if arguments.get("bg_color"):
                _set_cell_background(cell, arguments["bg_color"])
            align_str = arguments.get("alignment")
            for para in cell.paragraphs:
                if align_str:
                    para.alignment = ALIGN_MAP.get(align_str.upper(), WD_ALIGN_PARAGRAPH.LEFT)
                for run in para.runs:
                    _set_run_formatting(
                        run,
                        bold=arguments.get("bold"),
                        italic=arguments.get("italic"),
                        font_size=arguments.get("font_size"),
                        color_hex=arguments.get("color_hex"),
                    )
            _save()
            return ok(f"Cell [{row},{col}] in table {ti} updated.")

        elif name == "add_table_row":
            doc = _require_doc()
            ti = arguments["table_index"]
            data = arguments.get("data", [])
            tbl = doc.tables[ti]
            row = tbl.add_row()
            for ci, cell_text in enumerate(data):
                if ci < len(row.cells):
                    row.cells[ci].text = str(cell_text)
            _save()
            return ok(f"Row added to table {ti}.")

        elif name == "merge_table_cells":
            doc = _require_doc()
            ti = arguments["table_index"]
            tbl = doc.tables[ti]
            direction = arguments["direction"]
            if direction == "horizontal":
                row = arguments["row"]
                start_col = arguments["start_col"]
                end_col = arguments["end_col"]
                _merge_cells_horizontal(tbl, row, start_col, end_col)
            else:
                col = arguments["col"]
                start_row = arguments["start_row"]
                end_row = arguments["end_row"]
                _merge_cells_vertical(tbl, col, start_row, end_row)
            _save()
            return ok(f"Cells merged ({direction}) in table {ti}.")

        elif name == "format_table_cell":
            doc = _require_doc()
            ti = arguments["table_index"]
            row = arguments["row"]
            col = arguments["col"]
            cell = doc.tables[ti].cell(row, col)
            if arguments.get("bg_color"):
                _set_cell_background(cell, arguments["bg_color"])
            border_args = {}
            for side in ["top", "bottom", "left", "right"]:
                val = arguments.get(f"border_{side}")
                if val:
                    border_args[side] = val
            if border_args:
                _set_cell_border(
                    cell,
                    color=arguments.get("border_color", "000000"),
                    size=arguments.get("border_size", 4),
                    **border_args
                )
            _save()
            return ok(f"Cell [{row},{col}] in table {ti} formatted.")

        elif name == "delete_table_row":
            doc = _require_doc()
            ti = arguments["table_index"]
            row_idx = arguments["row"]
            tbl = doc.tables[ti]
            row = tbl.rows[row_idx]
            row._tr.getparent().remove(row._tr)
            _save()
            return ok(f"Row {row_idx} deleted from table {ti}.")

        elif name == "delete_table":
            doc = _require_doc()
            ti = arguments["table_index"]
            tbl = doc.tables[ti]
            tbl._tbl.getparent().remove(tbl._tbl)
            _save()
            return ok(f"Table {ti} deleted.")

        # ====================================================================
        # Headers & Footers
        # ====================================================================

        elif name == "set_header":
            doc = _require_doc()
            text = arguments.get("text", "")
            si = arguments.get("section_index", 0)
            which = arguments.get("which", "header")
            alignment = arguments.get("alignment", "LEFT")
            section = doc.sections[si]
            hf = _get_or_create_hdrftr(section, which)
            # Clear existing content
            for p in hf.paragraphs:
                for run in p.runs:
                    run.text = ""
            para = hf.paragraphs[0] if hf.paragraphs else hf.add_paragraph()
            para.clear()
            para.alignment = ALIGN_MAP.get(alignment.upper(), WD_ALIGN_PARAGRAPH.LEFT)
            run = para.add_run(text)
            _set_run_formatting(
                run,
                bold=arguments.get("bold"),
                italic=arguments.get("italic"),
                font_size=arguments.get("font_size"),
                font_name=arguments.get("font_name"),
            )
            _save()
            return ok(f"Header ({which}) set for section {si}.")

        elif name == "set_footer":
            doc = _require_doc()
            text = arguments.get("text", "")
            si = arguments.get("section_index", 0)
            which = arguments.get("which", "footer")
            alignment = arguments.get("alignment", "CENTER")
            add_page_number = arguments.get("add_page_number", False)
            page_number_format = arguments.get("page_number_format", "PAGE_NUMBER")
            section = doc.sections[si]
            hf = _get_or_create_hdrftr(section, which)
            para = hf.paragraphs[0] if hf.paragraphs else hf.add_paragraph()
            para.clear()
            para.alignment = ALIGN_MAP.get(alignment.upper(), WD_ALIGN_PARAGRAPH.CENTER)

            if text:
                run = para.add_run(text + ("  " if add_page_number else ""))
                _set_run_formatting(
                    run,
                    bold=arguments.get("bold"),
                    font_size=arguments.get("font_size"),
                    font_name=arguments.get("font_name"),
                )

            if add_page_number:
                _add_page_number_field(para, alignment=alignment, fmt=page_number_format)

            _save()
            return ok(f"Footer ({which}) set for section {si}.")

        elif name == "add_image_to_header":
            doc = _require_doc()
            image_path = arguments["image_path"]
            si = arguments.get("section_index", 0)
            width_cm = arguments.get("width_cm", 4.0)
            alignment = arguments.get("alignment", "LEFT")
            section = doc.sections[si]
            hf = _get_or_create_hdrftr(section, "header")
            para = hf.paragraphs[0] if hf.paragraphs else hf.add_paragraph()
            para.clear()
            para.alignment = ALIGN_MAP.get(alignment.upper(), WD_ALIGN_PARAGRAPH.LEFT)
            run = para.add_run()
            run.add_picture(image_path, width=Cm(width_cm))
            _save()
            return ok(f"Image added to header of section {si}.")

        elif name == "clear_header":
            doc = _require_doc()
            si = arguments.get("section_index", 0)
            which = arguments.get("which", "header")
            section = doc.sections[si]
            hf = getattr(section, which)
            for p in hf.paragraphs:
                for run in p.runs:
                    run.text = ""
            _save()
            return ok(f"Header ({which}) cleared for section {si}.")

        elif name == "clear_footer":
            doc = _require_doc()
            si = arguments.get("section_index", 0)
            which = arguments.get("which", "footer")
            section = doc.sections[si]
            hf = getattr(section, which)
            for p in hf.paragraphs:
                for run in p.runs:
                    run.text = ""
            _save()
            return ok(f"Footer ({which}) cleared for section {si}.")

        # ====================================================================
        # Images
        # ====================================================================

        elif name == "insert_image":
            doc = _require_doc()
            image_path = arguments["image_path"]
            width_cm = arguments.get("width_cm")
            height_cm = arguments.get("height_cm")
            alignment = arguments.get("alignment", "LEFT")
            caption_text = arguments.get("caption")
            after_text = arguments.get("after_text")
            after_heading = arguments.get("after_heading")

            # Build the image paragraph
            img_para = doc.add_paragraph()
            img_para.alignment = ALIGN_MAP.get(alignment.upper(), WD_ALIGN_PARAGRAPH.LEFT)
            run = img_para.add_run()

            kwargs = {}
            if width_cm:
                kwargs["width"] = Cm(width_cm)
            if height_cm:
                kwargs["height"] = Cm(height_cm)
            run.add_picture(image_path, **kwargs)

            # Reposition if needed (currently appended at end by python-docx API;
            # for precise positioning the XML element is moved)
            if after_text or after_heading:
                ref_para = None
                if after_heading:
                    ref_para = _find_heading_para(doc, after_heading)
                elif after_text:
                    for p in doc.paragraphs:
                        if after_text.lower() in p.text.lower():
                            ref_para = p
                            break
                if ref_para:
                    # Move the appended paragraph after ref_para
                    img_p_el = img_para._p
                    img_p_el.getparent().remove(img_p_el)
                    ref_para._p.addnext(img_p_el)

            if caption_text:
                cap_para = doc.add_paragraph(caption_text)
                try:
                    cap_para.style = doc.styles["Caption"]
                except Exception:
                    cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # Move caption after image
                cap_el = cap_para._p
                cap_el.getparent().remove(cap_el)
                img_para._p.addnext(cap_el)

            _save()
            return ok(f"Image '{image_path}' inserted.")

        # ====================================================================
        # Page layout
        # ====================================================================

        elif name == "set_page_margins":
            doc = _require_doc()
            si = arguments.get("section_index")
            sections = [doc.sections[si]] if si is not None else doc.sections
            for section in sections:
                if "top" in arguments:
                    section.top_margin = Cm(arguments["top"])
                if "bottom" in arguments:
                    section.bottom_margin = Cm(arguments["bottom"])
                if "left" in arguments:
                    section.left_margin = Cm(arguments["left"])
                if "right" in arguments:
                    section.right_margin = Cm(arguments["right"])
            _save()
            return ok("Page margins updated.")

        elif name == "set_page_orientation":
            doc = _require_doc()
            si = arguments.get("section_index", 0)
            orientation = arguments["orientation"].lower()
            section = doc.sections[si]
            if orientation == "landscape":
                section.orientation = WD_ORIENT.LANDSCAPE
                # Swap width/height if not already landscape
                if section.page_width < section.page_height:
                    section.page_width, section.page_height = (
                        section.page_height, section.page_width
                    )
            else:
                section.orientation = WD_ORIENT.PORTRAIT
                if section.page_width > section.page_height:
                    section.page_width, section.page_height = (
                        section.page_height, section.page_width
                    )
            _save()
            return ok(f"Section {si} orientation set to {orientation}.")

        elif name == "set_page_size":
            doc = _require_doc()
            si = arguments.get("section_index", 0)
            section = doc.sections[si]
            preset = arguments.get("preset")
            # Dimensions in EMU (1 cm = 360000 EMU)
            presets = {
                "A3": (Cm(29.7), Cm(42.0)),
                "A4": (Cm(21.0), Cm(29.7)),
                "A5": (Cm(14.8), Cm(21.0)),
                "Letter": (Cm(21.59), Cm(27.94)),
                "Legal": (Cm(21.59), Cm(35.56)),
            }
            if preset and preset in presets:
                w, h = presets[preset]
            elif "width_cm" in arguments and "height_cm" in arguments:
                w = Cm(arguments["width_cm"])
                h = Cm(arguments["height_cm"])
            else:
                return ok("Specify 'preset' or both 'width_cm' and 'height_cm'.")
            section.page_width = w
            section.page_height = h
            _save()
            size_label = preset or f"{arguments.get('width_cm')}x{arguments.get('height_cm')} cm"
            return ok(f"Page size set to {size_label}.")

        elif name == "add_section_break":
            doc = _require_doc()
            break_type = arguments.get("break_type", "new_page")
            break_map = {
                "new_page": WD_SECTION.NEW_PAGE,
                "continuous": WD_SECTION.CONTINUOUS,
                "even_page": WD_SECTION.EVEN_PAGE,
                "odd_page": WD_SECTION.ODD_PAGE,
            }
            after_text = arguments.get("after_text")
            ref_para = None
            if after_text:
                for p in doc.paragraphs:
                    if after_text.lower() in p.text.lower():
                        ref_para = p
                        break
            if ref_para:
                # Add section break via sectPr in the paragraph
                pPr = ref_para._p.get_or_add_pPr()
                sectPr = OxmlElement("w:sectPr")
                pgSz = OxmlElement("w:type")
                pgSz.set(qn("w:val"), list(break_map.keys())[
                    list(break_map.values()).index(break_map[break_type])
                ])
                sectPr.append(pgSz)
                pPr.append(sectPr)
            else:
                doc.add_section(break_map.get(break_type, WD_SECTION.NEW_PAGE))
            _save()
            return ok(f"Section break ({break_type}) inserted.")

        elif name == "add_page_break":
            doc = _require_doc()
            after_text = arguments.get("after_text")
            at_index = arguments.get("at_index")
            at_end = arguments.get("at_end", False)

            ref = None
            if after_text:
                for p in doc.paragraphs:
                    if after_text.lower() in p.text.lower():
                        ref = p
                        break
            elif at_index is not None:
                ref = doc.paragraphs[at_index]

            if ref:
                run = ref.add_run()
                run.add_break(WD_BREAK.PAGE)
            else:
                doc.add_page_break()
            _save()
            return ok("Page break inserted.")

        elif name == "set_columns":
            doc = _require_doc()
            si = arguments.get("section_index", 0)
            num_cols = arguments["num_cols"]
            spacing_cm = arguments.get("spacing_cm", 1.25)
            _set_section_columns(doc.sections[si], num_cols, spacing_cm)
            _save()
            return ok(f"Section {si} set to {num_cols} column(s).")

        # ====================================================================
        # Styles
        # ====================================================================

        elif name == "create_style":
            doc = _require_doc()
            style = _create_paragraph_style(
                doc,
                style_name=arguments["style_name"],
                base_style=arguments.get("base_style", "Normal"),
                font_name=arguments.get("font_name"),
                font_size=arguments.get("font_size"),
                bold=arguments.get("bold"),
                italic=arguments.get("italic"),
                color_hex=arguments.get("color_hex"),
                alignment=arguments.get("alignment"),
                space_before=arguments.get("space_before"),
                space_after=arguments.get("space_after"),
            )
            _save()
            return ok(f"Style '{style.name}' created/updated.")

        # ====================================================================
        # TOC
        # ====================================================================

        elif name == "insert_toc":
            doc = _require_doc()
            title = arguments.get("title", "Table of Contents")
            max_level = arguments.get("max_level", 3)
            after_heading = arguments.get("after_heading")
            at_end = arguments.get("at_end", False)
            _insert_toc(doc, title=title, max_level=max_level,
                        after_heading=after_heading if not at_end else "__end__")
            _save()
            return ok(
                f"TOC inserted (levels 1-{max_level}). "
                "Open in Word and press Ctrl+A then F9 to render the TOC."
            )

        # ====================================================================
        # Hyperlinks
        # ====================================================================

        elif name == "add_hyperlink":
            doc = _require_doc()
            para_match = arguments.get("paragraph_match", "")
            para_idx = arguments.get("paragraph_index")
            link_text = arguments["text"]
            url = arguments["url"]
            color_hex = arguments.get("color_hex", "0563C1")
            underline = arguments.get("underline", True)

            target = None
            if para_idx is not None:
                target = doc.paragraphs[para_idx]
            elif para_match:
                for p in doc.paragraphs:
                    if para_match.lower() in p.text.lower():
                        target = p
                        break
            if target is None:
                return ok("Paragraph not found.")
            _add_hyperlink(target, link_text, url, color_hex=color_hex, underline=underline)
            _save()
            return ok(f"Hyperlink '{link_text}' -> '{url}' added.")

        # ====================================================================
        # Comments
        # ====================================================================

        elif name == "add_comment":
            doc = _require_doc()
            match_text = arguments["match_text"]
            comment_text = arguments["comment"]
            author = arguments.get("author", "Claude")
            initials = arguments.get("initials", "AI")
            for p in doc.paragraphs:
                if match_text.lower() in p.text.lower():
                    cid = _add_native_comment(doc, p, comment_text,
                                              author=author, initials=initials)
                    _save()
                    return ok(f"Comment (id={cid}) added to paragraph containing '{match_text}'.")
            return ok(f"Text '{match_text}' not found.")

        # ====================================================================
        # Bookmarks
        # ====================================================================

        elif name == "add_bookmark":
            doc = _require_doc()
            para_idx = arguments.get("paragraph_index")
            match_text = arguments.get("match_text", "")
            bookmark_name = arguments["bookmark_name"]
            bookmark_id = arguments.get("bookmark_id", 1)

            target = None
            if para_idx is not None:
                target = doc.paragraphs[para_idx]
            elif match_text:
                for p in doc.paragraphs:
                    if match_text.lower() in p.text.lower():
                        target = p
                        break
            if target is None:
                return ok("Paragraph not found.")
            _add_bookmark(target, bookmark_name, bookmark_id)
            _save()
            return ok(f"Bookmark '{bookmark_name}' added.")

        # ====================================================================
        # Document properties
        # ====================================================================

        elif name == "set_document_properties":
            doc = _require_doc()
            cp = doc.core_properties
            if "title" in arguments:
                cp.title = arguments["title"]
            if "author" in arguments:
                cp.author = arguments["author"]
            if "subject" in arguments:
                cp.subject = arguments["subject"]
            if "keywords" in arguments:
                cp.keywords = arguments["keywords"]
            if "description" in arguments:
                cp.description = arguments["description"]
            if "category" in arguments:
                cp.category = arguments["category"]
            _save()
            return ok("Document properties updated.")

        # ====================================================================
        # Advanced XML
        # ====================================================================

        elif name == "apply_xml_patch":
            doc = _require_doc()
            idx = arguments["paragraph_index"]
            xml_content = arguments["xml_content"]
            old_p = doc.paragraphs[idx]._p
            new_p = etree.fromstring(xml_content.encode())
            old_p.getparent().replace(old_p, new_p)
            _save()
            return ok(f"XML patch applied to paragraph {idx}.")

        # ====================================================================
        # Batch document builder
        # ====================================================================

        elif name == "build_document":
            path = arguments["path"]
            template = arguments.get("template")
            doc = Document(template) if template else Document()

            # Default font/size
            if arguments.get("default_font") or arguments.get("default_font_size"):
                normal = doc.styles["Normal"]
                if arguments.get("default_font"):
                    normal.font.name = arguments["default_font"]
                if arguments.get("default_font_size"):
                    normal.font.size = Pt(arguments["default_font_size"])

            # Document properties
            cp = doc.core_properties
            for prop in ["title", "author", "subject"]:
                if arguments.get(prop):
                    setattr(cp, prop, arguments[prop])

            # Margins
            if arguments.get("margins"):
                m = arguments["margins"]
                for section in doc.sections:
                    if "top" in m:
                        section.top_margin = Cm(m["top"])
                    if "bottom" in m:
                        section.bottom_margin = Cm(m["bottom"])
                    if "left" in m:
                        section.left_margin = Cm(m["left"])
                    if "right" in m:
                        section.right_margin = Cm(m["right"])

            def apply_run_fmt(run, data: dict):
                _set_run_formatting(
                    run,
                    bold=data.get("bold"),
                    italic=data.get("italic"),
                    underline=data.get("underline"),
                    strike=data.get("strike"),
                    font_size=data.get("font_size"),
                    font_name=data.get("font_name"),
                    color_hex=data.get("color_hex"),
                )

            def apply_para_format(para, el: dict):
                if el.get("alignment"):
                    para.alignment = ALIGN_MAP.get(el["alignment"].upper(), WD_ALIGN_PARAGRAPH.LEFT)
                pf = para.paragraph_format
                if el.get("space_before") is not None:
                    pf.space_before = Pt(el["space_before"])
                if el.get("space_after") is not None:
                    pf.space_after = Pt(el["space_after"])

            def fill_paragraph(para, el: dict):
                """Fill paragraph with text or runs (with optional hyperlinks)."""
                if el.get("runs"):
                    for rd in el["runs"]:
                        if rd.get("url"):
                            _add_hyperlink(para, rd.get("text", ""), rd["url"],
                                           color_hex=rd.get("color_hex", "0563C1"))
                        else:
                            run = para.add_run(rd.get("text", ""))
                            apply_run_fmt(run, rd)
                elif el.get("text"):
                    para.add_run(el["text"])

            added = 0
            for el in arguments.get("elements", []):
                el_type = el.get("type", "paragraph")

                if el_type == "page_break":
                    doc.add_page_break()

                elif el_type == "section_break":
                    break_map = {
                        "new_page": WD_SECTION.NEW_PAGE,
                        "continuous": WD_SECTION.CONTINUOUS,
                        "even_page": WD_SECTION.EVEN_PAGE,
                        "odd_page": WD_SECTION.ODD_PAGE,
                    }
                    bt = el.get("break_type", "new_page")
                    doc.add_section(break_map.get(bt, WD_SECTION.NEW_PAGE))

                elif el_type == "toc":
                    _insert_toc(doc,
                                title=el.get("toc_title", "Table of Contents"),
                                max_level=el.get("toc_max_level", 3))

                elif el_type == "heading":
                    level = el.get("level", 1)
                    para = doc.add_heading("", level=level)
                    fill_paragraph(para, el)
                    apply_para_format(para, el)

                elif el_type in ("paragraph", "list"):
                    list_type = el.get("list_type")
                    list_level = el.get("list_level", 0)
                    style_name = el.get("style")
                    if list_type == "bullet":
                        style_name = style_name or (
                            "List Bullet" if list_level == 0
                            else f"List Bullet {list_level + 1}"
                        )
                    elif list_type == "numbered":
                        style_name = style_name or (
                            "List Number" if list_level == 0
                            else f"List Number {list_level + 1}"
                        )
                    else:
                        style_name = style_name or "Normal"
                    try:
                        para = doc.add_paragraph(style=style_name)
                    except Exception:
                        para = doc.add_paragraph()
                    fill_paragraph(para, el)
                    apply_para_format(para, el)

                elif el_type == "image":
                    image_path = el.get("image_path", "")
                    if image_path and os.path.exists(image_path):
                        img_para = doc.add_paragraph()
                        alignment = el.get("alignment", "LEFT")
                        img_para.alignment = ALIGN_MAP.get(alignment.upper(), WD_ALIGN_PARAGRAPH.LEFT)
                        run = img_para.add_run()
                        kwargs = {}
                        if el.get("width_cm"):
                            kwargs["width"] = Cm(el["width_cm"])
                        run.add_picture(image_path, **kwargs)
                        if el.get("caption"):
                            cap = doc.add_paragraph(el["caption"])
                            try:
                                cap.style = doc.styles["Caption"]
                            except Exception:
                                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

                elif el_type == "table":
                    data = el.get("data", [])
                    if not data:
                        added += 1
                        continue
                    rows = len(data)
                    cols = max(len(r) for r in data) if data else 1
                    table_style = el.get("table_style", "Table Grid")
                    tbl = doc.add_table(rows=rows, cols=cols)
                    try:
                        tbl.style = doc.styles[table_style]
                    except Exception:
                        pass
                    col_widths = el.get("col_widths", [])
                    header_row = el.get("header_row", False)
                    for ri, row_data in enumerate(data):
                        for ci, cell_val in enumerate(row_data):
                            if ci >= cols:
                                break
                            cell = tbl.cell(ri, ci)
                            if isinstance(cell_val, dict):
                                cell.text = ""
                                run = cell.paragraphs[0].add_run(
                                    str(cell_val.get("text", ""))
                                )
                                apply_run_fmt(run, cell_val)
                                if cell_val.get("bg_color"):
                                    _set_cell_background(cell, cell_val["bg_color"])
                            else:
                                cell.text = str(cell_val)
                            if header_row and ri == 0:
                                for para in cell.paragraphs:
                                    for run in para.runs:
                                        run.bold = True
                    if col_widths:
                        for ci, width_cm in enumerate(col_widths):
                            if ci < cols:
                                for row in tbl.rows:
                                    row.cells[ci].width = Cm(width_cm)

                elif el_type == "hyperlink":
                    para = doc.add_paragraph()
                    style_name = el.get("style", "Normal")
                    try:
                        para.style = doc.styles[style_name]
                    except Exception:
                        pass
                    apply_para_format(para, el)
                    _add_hyperlink(
                        para,
                        text=el.get("text", el.get("url", "")),
                        url=el.get("url", ""),
                        color_hex=el.get("color_hex", "0563C1"),
                    )

                added += 1

            # Global header
            if arguments.get("header_text"):
                hdr = doc.sections[0].header
                p = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
                p.clear()
                p.add_run(arguments["header_text"])

            # Global footer
            if arguments.get("footer_text") or arguments.get("footer_page_numbers"):
                ftr = doc.sections[0].footer
                p = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
                p.clear()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if arguments.get("footer_text"):
                    p.add_run(arguments["footer_text"])
                if arguments.get("footer_page_numbers"):
                    _add_page_number_field(p, alignment="CENTER", fmt="PAGE_OF_PAGES")

            doc.save(path)
            _current_doc_path = path
            _doc = Document(path)
            _reload_in_word(path)
            return ok(
                f"Document built: {path}\n"
                f"Elements: {added} | "
                f"Paragraphs: {len(_doc.paragraphs)} | "
                f"Tables: {len(_doc.tables)}"
            )

        # ====================================================================
        # Tracked changes (adapted from knorq-ai/docx-mcp-server)
        # ====================================================================

        elif name == "accept_all_changes":
            doc = _require_doc()
            _accept_all_changes(doc)
            _save()
            return ok("All tracked changes accepted and document saved.")

        elif name == "reject_all_changes":
            doc = _require_doc()
            _reject_all_changes(doc)
            _save()
            return ok("All tracked changes rejected and document saved.")

        elif name == "has_tracked_changes":
            doc = _require_doc()
            has = _has_tracked_changes(doc)
            return ok(f"Document has tracked changes: {has}")

        # ====================================================================
        # Comment extraction (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================

        elif name == "get_all_comments":
            doc = _require_doc()
            comments = _extract_all_comments(doc)
            return ok(json.dumps({
                "success": True,
                "comments": comments,
                "total_comments": len(comments)
            }, indent=2))

        elif name == "get_comments_by_author":
            doc = _require_doc()
            author = arguments["author"]
            comments = _extract_all_comments(doc)
            filtered = [c for c in comments if c.get("author", "").lower() == author.lower()]
            return ok(json.dumps({
                "success": True,
                "comments": filtered,
                "total_comments": len(filtered),
                "filter_author": author
            }, indent=2))

        # ====================================================================
        # Footnotes & Endnotes (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================

        elif name == "add_footnote":
            doc = _require_doc()
            para_idx = arguments["paragraph_index"]
            fn_text = arguments["footnote_text"]
            paras = doc.paragraphs
            if para_idx < 1 or para_idx > len(paras):
                return ok(f"ERROR: Invalid paragraph index. Document has {len(paras)} paragraphs (1-{len(paras)}).")
            para = paras[para_idx - 1]
            fn_id = _add_footnote_to_paragraph(para, fn_text)
            _save()
            return ok(f"Footnote (id={fn_id}) added to paragraph {para_idx}.")

        elif name == "read_footnotes":
            doc = _require_doc()
            footnotes = _read_footnotes(doc)
            return ok(json.dumps({
                "success": True,
                "footnotes": footnotes,
                "total_footnotes": len(footnotes)
            }, indent=2))

        elif name == "add_endnote":
            doc = _require_doc()
            para_idx = arguments["paragraph_index"]
            en_text = arguments["endnote_text"]
            paras = doc.paragraphs
            if para_idx < 1 or para_idx > len(paras):
                return ok(f"ERROR: Invalid paragraph index. Document has {len(paras)} paragraphs (1-{len(paras)}).")
            para = paras[para_idx - 1]
            # Add endnote marker (superscript dagger)
            run = para.add_run("\u2020")
            run.font.superscript = True
            # Add endnotes section if not present
            has_endnotes = any(p.text.strip() in ("Endnotes:", "ENDNOTES") for p in doc.paragraphs)
            if not has_endnotes:
                doc.add_page_break()
                doc.add_heading("Endnotes:", level=1)
            en_para = doc.add_paragraph(f"\u2020 {en_text}")
            try:
                en_para.style = doc.styles["Endnote Text"]
            except KeyError:
                pass
            _save()
            return ok(f"Endnote added to paragraph {para_idx}.")

        # ====================================================================
        # Advanced table formatting (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================

        elif name == "format_table":
            doc = _require_doc()
            tbl_idx = arguments["table_index"]
            if tbl_idx < 1 or tbl_idx > len(doc.tables):
                return ok(f"ERROR: Invalid table index. Document has {len(doc.tables)} tables (1-{len(doc.tables)}).")
            table = doc.tables[tbl_idx - 1]
            has_header = arguments.get("has_header_row", False)
            border_style = arguments.get("border_style")
            shading = arguments.get("shading")
            # Apply header row formatting
            if has_header and table.rows:
                for cell in table.rows[0].cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
            # Apply borders
            if border_style:
                val_map = {"none": "nil", "single": "single", "double": "double", "thick": "thick"}
                val = val_map.get(border_style.lower(), "single")
                for row in table.rows:
                    for cell in row.cells:
                        _set_cell_border(cell, top=val, bottom=val, left=val, right=val, color="000000")
            # Apply cell shading
            if shading:
                for i, row_colors in enumerate(shading):
                    if i >= len(table.rows):
                        break
                    for j, color in enumerate(row_colors):
                        if j >= len(table.rows[i].cells):
                            break
                        _set_cell_background(table.rows[i].cells[j], color)
            _save()
            return ok(f"Table {tbl_idx} formatted successfully.")

        elif name == "set_table_cell_shading":
            doc = _require_doc()
            tbl_idx = arguments["table_index"]
            row_idx = arguments["row_index"]
            col_idx = arguments["col_index"]
            fill_color = arguments["fill_color"]
            if tbl_idx < 1 or tbl_idx > len(doc.tables):
                return ok(f"ERROR: Invalid table index.")
            table = doc.tables[tbl_idx - 1]
            if row_idx < 1 or row_idx > len(table.rows):
                return ok(f"ERROR: Invalid row index.")
            if col_idx < 1 or col_idx > len(table.rows[row_idx - 1].cells):
                return ok(f"ERROR: Invalid column index.")
            cell = table.rows[row_idx - 1].cells[col_idx - 1]
            _set_cell_background(cell, fill_color)
            _save()
            return ok(f"Cell shading applied to table {tbl_idx}, row {row_idx}, col {col_idx}.")

        elif name == "apply_table_alternating_rows":
            doc = _require_doc()
            tbl_idx = arguments["table_index"]
            color1 = arguments.get("color1", "FFFFFF")
            color2 = arguments.get("color2", "F2F2F2")
            if tbl_idx < 1 or tbl_idx > len(doc.tables):
                return ok(f"ERROR: Invalid table index.")
            table = doc.tables[tbl_idx - 1]
            for i, row in enumerate(table.rows):
                color = color1 if i % 2 == 0 else color2
                for cell in row.cells:
                    _set_cell_background(cell, color)
            _save()
            return ok(f"Alternating row shading applied to table {tbl_idx}.")

        elif name == "highlight_table_header":
            doc = _require_doc()
            tbl_idx = arguments["table_index"]
            header_color = arguments.get("header_color", "4472C4")
            text_color = arguments.get("text_color", "FFFFFF")
            if tbl_idx < 1 or tbl_idx > len(doc.tables):
                return ok(f"ERROR: Invalid table index.")
            table = doc.tables[tbl_idx - 1]
            if table.rows:
                for cell in table.rows[0].cells:
                    _set_cell_background(cell, header_color)
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                            h = text_color.lstrip("#")
                            run.font.color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
            _save()
            return ok(f"Header highlighting applied to table {tbl_idx}.")

        elif name == "set_table_cell_alignment":
            doc = _require_doc()
            tbl_idx = arguments["table_index"]
            row_idx = arguments["row_index"]
            col_idx = arguments["col_index"]
            horizontal = arguments.get("horizontal", "left")
            vertical = arguments.get("vertical", "top")
            if tbl_idx < 1 or tbl_idx > len(doc.tables):
                return ok(f"ERROR: Invalid table index.")
            table = doc.tables[tbl_idx - 1]
            if row_idx < 1 or row_idx > len(table.rows):
                return ok(f"ERROR: Invalid row index.")
            if col_idx < 1 or col_idx > len(table.rows[row_idx - 1].cells):
                return ok(f"ERROR: Invalid column index.")
            cell = table.rows[row_idx - 1].cells[col_idx - 1]
            # Set horizontal alignment
            align_map = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                        "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
            for para in cell.paragraphs:
                para.alignment = align_map.get(horizontal.lower(), WD_ALIGN_PARAGRAPH.LEFT)
            # Set vertical alignment
            v_align_map = {"top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
                          "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
                          "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM}
            cell.vertical_alignment = v_align_map.get(vertical.lower(), WD_CELL_VERTICAL_ALIGNMENT.TOP)
            _save()
            return ok(f"Cell alignment set for table {tbl_idx}, cell ({row_idx},{col_idx}) to {horizontal}/{vertical}.")

        elif name == "set_table_column_width":
            doc = _require_doc()
            tbl_idx = arguments["table_index"]
            col_idx = arguments["col_index"]
            width = arguments["width"]
            width_type = arguments.get("width_type", "points")
            if tbl_idx < 1 or tbl_idx > len(doc.tables):
                return ok(f"ERROR: Invalid table index.")
            table = doc.tables[tbl_idx - 1]
            if col_idx < 1 or col_idx > len(table.columns):
                return ok(f"ERROR: Invalid column index.")
            # Convert width to EMU
            if width_type == "inches":
                width_val = Inches(width)
            elif width_type == "cm":
                width_val = Cm(width)
            elif width_type == "points":
                width_val = Pt(width)
            else:
                width_val = Inches(width / 72.0)  # fallback: treat as points
            for row in table.rows:
                if col_idx - 1 < len(row.cells):
                    row.cells[col_idx - 1].width = width_val
            _save()
            return ok(f"Column {col_idx} width set to {width} {width_type} in table {tbl_idx}.")

        elif name == "set_table_column_widths":
            doc = _require_doc()
            tbl_idx = arguments["table_index"]
            widths = arguments["widths"]
            width_type = arguments.get("width_type", "points")
            if tbl_idx < 1 or tbl_idx > len(doc.tables):
                return ok(f"ERROR: Invalid table index.")
            table = doc.tables[tbl_idx - 1]
            for ci, width in enumerate(widths):
                if ci >= len(table.columns):
                    break
                if width_type == "inches":
                    width_val = Inches(width)
                elif width_type == "cm":
                    width_val = Cm(width)
                elif width_type == "points":
                    width_val = Pt(width)
                else:
                    width_val = Inches(width / 72.0)
                for row in table.rows:
                    if ci < len(row.cells):
                        row.cells[ci].width = width_val
            _save()
            return ok(f"Column widths set for {len(widths)} columns in table {tbl_idx}.")

        elif name == "set_table_cell_padding":
            doc = _require_doc()
            tbl_idx = arguments["table_index"]
            row_idx = arguments["row_index"]
            col_idx = arguments["col_index"]
            if tbl_idx < 1 or tbl_idx > len(doc.tables):
                return ok(f"ERROR: Invalid table index.")
            table = doc.tables[tbl_idx - 1]
            if row_idx < 1 or row_idx > len(table.rows):
                return ok(f"ERROR: Invalid row index.")
            if col_idx < 1 or col_idx > len(table.rows[row_idx - 1].cells):
                return ok(f"ERROR: Invalid column index.")
            cell = table.rows[row_idx - 1].cells[col_idx - 1]
            _set_cell_padding(cell,
                             top=arguments.get("top"),
                             bottom=arguments.get("bottom"),
                             left=arguments.get("left"),
                             right=arguments.get("right"))
            _save()
            return ok(f"Cell padding set for table {tbl_idx}, cell ({row_idx},{col_idx}).")

        # ====================================================================
        # Standalone list tools (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================

        elif name == "insert_bulleted_list":
            doc = _require_doc()
            items = arguments["items"]
            style = arguments.get("style", "List Bullet")
            result = _insert_bulleted_list(doc, items, style)
            _save()
            return ok(result)

        elif name == "insert_numbered_list":
            doc = _require_doc()
            items = arguments["items"]
            style = arguments.get("style", "List Number")
            result = _insert_numbered_list(doc, items, style)
            _save()
            return ok(result)

        # ====================================================================
        # Text highlighting (adapted from knorq-ai/docx-mcp-server)
        # ====================================================================

        elif name == "highlight_text":
            doc = _require_doc()
            search_text = arguments["search_text"]
            color = arguments.get("color", "yellow")
            match_case = arguments.get("match_case", False)
            total = 0
            for para in doc.paragraphs:
                total += _highlight_text_in_paragraph(para, search_text, color, match_case)
            # Also search in table cells
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            total += _highlight_text_in_paragraph(para, search_text, color, match_case)
            if total > 0:
                _save()
                return ok(f"Highlighted {total} occurrence(s) of '{search_text}' with {color}.")
            else:
                return ok(f"No occurrences of '{search_text}' found.")

        # ====================================================================
        # Document protection (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================

        elif name == "protect_document":
            doc = _require_doc()
            password = arguments["password"]
            # Save first to ensure latest content is on disk
            _save()
            result = _protect_document_with_password(_current_doc_path, password)
            return ok(result)

        elif name == "unprotect_document":
            doc = _require_doc()
            password = arguments["password"]
            result = _unprotect_document_with_password(_current_doc_path, password)
            if "successfully" in result:
                # Reload the decrypted document
                _doc = Document(_current_doc_path)
            return ok(result)

        # ====================================================================
        # PDF export & Document merge (adapted from GongRzhe/Office-Word-MCP-Server)
        # ====================================================================

        elif name == "convert_to_pdf":
            doc = _require_doc()
            output_path = arguments.get("output_path")
            # Save first to ensure latest content is on disk
            _save()
            result = _convert_to_pdf(_current_doc_path, output_path)
            return ok(result)

        elif name == "merge_documents":
            source_paths = arguments["source_paths"]
            output_path = arguments["output_path"]
            result = _merge_documents(source_paths, output_path)
            return ok(result)

        # ====================================================================
        # Stable anchors, images, comment threading (adapted from knorq-ai)
        # ====================================================================

        elif name == "ensure_anchors":
            doc = _require_doc()
            result = _ensure_anchors(doc)
            _save()
            blocks_str = '\n'.join(
                f"  [{b['index']}] @{b['anchor'] or '(no anchor)'}"
                f"{' [table]' if b['type'] == 'table' else ''}"
                f"{' [sdt]' if b['type'] == 'sdt' else ''}"
                f" {b['text_preview']}"
                for b in result['blocks']
            )
            return ok(
                f"Anchors: seeded {result['seeded']}, repaired {result['repaired']}, "
                f"{len(result['blocks'])} block(s).\n\n{blocks_str}\n\n"
                f"<json>\n{json.dumps(result, indent=2)}\n</json>"
            )

        elif name == "list_images":
            doc = _require_doc()
            images = _list_images(doc)
            return ok(json.dumps({
                "success": True,
                "images": images,
                "total_images": len(images)
            }, indent=2))

        elif name == "reply_to_comment":
            doc = _require_doc()
            comment_id = arguments["comment_id"]
            reply_text = arguments["reply_text"]
            author = arguments.get("author", "Claude")
            result = _reply_to_comment(doc, comment_id, reply_text, author)
            _save()
            return ok(result)

        elif name == "delete_comment":
            doc = _require_doc()
            comment_id = arguments["comment_id"]
            result = _delete_comment(doc, comment_id)
            _save()
            return ok(result)

        elif name == "get_page_layout":
            doc = _require_doc()
            layout = _get_page_layout(doc)
            return ok(json.dumps(layout, indent=2, default=str))

        elif name == "format_text":
            doc = _require_doc()
            search = arguments["search"]
            count = _format_text_in_paragraphs(
                doc, search,
                bold=arguments.get("bold"),
                italic=arguments.get("italic"),
                underline=arguments.get("underline"),
                strikethrough=arguments.get("strikethrough"),
                highlight_color=arguments.get("highlight_color"),
                font_name=arguments.get("font_name"),
                font_size=arguments.get("font_size"),
                font_color=arguments.get("font_color"),
                match_case=arguments.get("match_case", False),
            )
            if count > 0:
                _save()
                return ok(f"Formatted {count} run(s) matching '{search}'.")
            else:
                return ok(f"No runs matching '{search}' found.")

        elif name == "set_headings":
            doc = _require_doc()
            headings = arguments["headings"]
            result = _set_headings_bulk(doc, headings)
            _save()
            return ok(result)

        elif name == "get_paragraph_format":
            doc = _require_doc()
            para_idx = arguments["paragraph_index"]
            result = _get_paragraph_format(doc, para_idx)
            return ok(json.dumps(result, indent=2, default=str))

        # ====================================================================
        # Bibliography / Sources
        # ====================================================================

        elif name == "add_bibliography_source":
            src = {
                "source_type": arguments.get("source_type", "JournalArticle"),
                "tag": arguments.get("tag", ""),
                "title": arguments.get("title", ""),
                "year": arguments.get("year", ""),
                "authors": arguments.get("authors", []),
                "journal_name": arguments.get("journal_name", ""),
                "book_title": arguments.get("book_title", ""),
                "volume": arguments.get("volume", ""),
                "issue": arguments.get("issue", ""),
                "pages": arguments.get("pages", ""),
                "publisher": arguments.get("publisher", ""),
                "city": arguments.get("city", ""),
                "doi": arguments.get("doi", ""),
                "url": arguments.get("url", ""),
                "abstract": arguments.get("abstract", ""),
                "edition": arguments.get("edition", ""),
                "isbn": arguments.get("isbn", ""),
                "issn": arguments.get("issn", ""),
                "language": arguments.get("language", ""),
                "comments": arguments.get("comments", ""),
                "conference_name": arguments.get("conference_name", ""),
                "medium": arguments.get("medium", ""),
                "number": arguments.get("number", ""),
                "country": arguments.get("country", ""),
                "institution": arguments.get("institution", ""),
                "department": arguments.get("department", ""),
                "thesis_type": arguments.get("thesis_type", ""),
                "access_date": arguments.get("access_date", ""),
                "production_company": arguments.get("production_company", ""),
                "guid": arguments.get("guid", ""),
            }
            added, skipped, messages = _add_sources_to_xml([src])
            result_lines = [f"Added {added} source(s), skipped {skipped} duplicate(s)."]
            result_lines.extend(messages)
            return ok("\n".join(result_lines))

        elif name == "add_bibliography_sources_batch":
            sources = arguments.get("sources", [])
            if not sources:
                return ok("No sources provided.")
            added, skipped, messages = _add_sources_to_xml(sources)
            result_lines = [f"Added {added} source(s), skipped {skipped} duplicate(s)."]
            result_lines.extend(messages)
            return ok("\n".join(result_lines))

        elif name == "list_bibliography_sources":
            sources = _list_sources_from_xml()
            if not sources:
                return ok("No sources found in the master bibliography.")
            lines = [f"Total sources: {len(sources)}\n"]
            for i, s in enumerate(sources, 1):
                lines.append(
                    f"  {i}. [{s['source_type']}] {s['first_author']} ({s['year']}) - {s['title']}"
                )
                lines.append(f"     Tag: {s['tag']}")
            return ok("\n".join(lines))

        elif name == "clear_bibliography_sources":
            count = _clear_sources_xml()
            return ok(f"Cleared {count} source(s) from the master bibliography.")

        elif name == "insert_citation":
            doc = _require_doc()
            tags = arguments.get("source_tags", [])
            if not tags:
                return ok("No source_tags provided.")
            match_text = arguments.get("match_text", "")
            para_idx = arguments.get("paragraph_index")
            tbl_idx = arguments.get("table_index")
            row = arguments.get("row")
            col = arguments.get("col")

            # Table cell mode
            if tbl_idx is not None and row is not None and col is not None:
                table = doc.tables[tbl_idx - 1]
                cell = table.rows[row - 1].cells[col - 1]
                _replace_cell_with_citations(cell, tags)
                _save()
                return ok(f"Replaced cell (table {tbl_idx}, row {row}, col {col}) with citation field(s) for tags: {', '.join(tags)}")

            # Paragraph mode
            replaced = False
            if para_idx is not None:
                # Search specific paragraph (1-based)
                paras = doc.paragraphs
                if 1 <= para_idx <= len(paras):
                    para = paras[para_idx - 1]
                    if match_text:
                        replaced = _replace_text_with_citation_in_para(para, match_text, tags)
                        if replaced:
                            _save()
                            return ok(f"Replaced '{match_text}' in paragraph {para_idx} with citation field(s) for tags: {', '.join(tags)}")
                        else:
                            return ok(f"Text '{match_text}' not found in paragraph {para_idx}.")
                    else:
                        # Append citation to end of paragraph
                        p_elem = para._element
                        for cf_run in _make_citation_field_runs(tags):
                            p_elem.append(cf_run)
                        _save()
                        return ok(f"Appended citation field(s) to paragraph {para_idx} for tags: {', '.join(tags)}")
                else:
                    return ok(f"Paragraph index {para_idx} out of range (1-{len(paras)}).")
            else:
                # Search all paragraphs
                if not match_text:
                    return ok("Either match_text or paragraph_index must be provided.")
                for i, para in enumerate(doc.paragraphs, 1):
                    if match_text in para.text:
                        replaced = _replace_text_with_citation_in_para(para, match_text, tags)
                        if replaced:
                            _save()
                            return ok(f"Replaced '{match_text}' in paragraph {i} with citation field(s) for tags: {', '.join(tags)}")
                if not replaced:
                    return ok(f"Text '{match_text}' not found in any paragraph.")

        elif name == "insert_citations_batch":
            doc = _require_doc()
            citations = arguments.get("citations", [])
            if not citations:
                return ok("No citations provided.")
            results = []
            for item in citations:
                tags = item.get("source_tags", [])
                if not tags:
                    results.append("SKIP: no source_tags")
                    continue
                match_text = item.get("match_text", "")
                para_idx = item.get("paragraph_index")
                tbl_idx = item.get("table_index")
                row = item.get("row")
                col = item.get("col")

                # Table cell mode
                if tbl_idx is not None and row is not None and col is not None:
                    if 1 <= tbl_idx <= len(doc.tables):
                        table = doc.tables[tbl_idx - 1]
                        if 1 <= row <= len(table.rows) and 1 <= col <= len(table.rows[row - 1].cells):
                            cell = table.rows[row - 1].cells[col - 1]
                            _replace_cell_with_citations(cell, tags)
                            results.append(f"OK: table {tbl_idx} row {row} col {col} -> {', '.join(tags)}")
                        else:
                            results.append(f"SKIP: table {tbl_idx} row {row} col {col} out of range")
                    else:
                        results.append(f"SKIP: table {tbl_idx} out of range")
                    continue

                # Paragraph mode
                replaced = False
                if para_idx is not None:
                    paras = doc.paragraphs
                    if 1 <= para_idx <= len(paras):
                        para = paras[para_idx - 1]
                        if match_text:
                            replaced = _replace_text_with_citation_in_para(para, match_text, tags)
                            if replaced:
                                results.append(f"OK: para {para_idx} '{match_text}' -> {', '.join(tags)}")
                            else:
                                results.append(f"SKIP: '{match_text}' not found in para {para_idx}")
                        else:
                            p_elem = para._element
                            for cf_run in _make_citation_field_runs(tags):
                                p_elem.append(cf_run)
                            results.append(f"OK: appended to para {para_idx} -> {', '.join(tags)}")
                    else:
                        results.append(f"SKIP: para {para_idx} out of range")
                else:
                    if not match_text:
                        results.append("SKIP: no match_text or paragraph_index")
                        continue
                    for i, para in enumerate(doc.paragraphs, 1):
                        if match_text in para.text:
                            replaced = _replace_text_with_citation_in_para(para, match_text, tags)
                            if replaced:
                                results.append(f"OK: para {i} '{match_text}' -> {', '.join(tags)}")
                                break
                    if not replaced:
                        results.append(f"SKIP: '{match_text}' not found in any paragraph")
            _save()
            succeeded = sum(1 for r in results if r.startswith("OK"))
            skipped = sum(1 for r in results if r.startswith("SKIP"))
            return ok(f"Batch complete: {succeeded} replaced, {skipped} skipped.\n" + "\n".join(results))

        elif name == "copy_sources_to_current_list":
            doc = _require_doc()
            if not _current_doc_path:
                return ok("No document path set.")
            # Save first to persist any pending changes
            _save()
            # Now manipulate the .docx zip to update customXml/item1.xml
            count = _copy_sources_to_current_list(_current_doc_path)
            # Reload the document in python-docx to pick up changes
            _doc = Document(_current_doc_path)
            return ok(f"Copied {count} source(s) from master bibliography to the document's Current List.")

        elif name == "insert_bibliography":
            doc = _require_doc()
            heading_text = arguments.get("heading_text", "References")
            heading_style = arguments.get("heading_style", "Heading 1")

            # Add heading paragraph
            bib_heading = doc.add_paragraph()
            try:
                bib_heading.style = doc.styles[heading_style]
            except KeyError:
                bib_heading.style = doc.styles['Heading 1']
            bib_heading.add_run(heading_text)

            # Add bibliography field paragraph
            bib_para = doc.add_paragraph()
            p_elem = bib_para._element
            for bf_run in _make_bibliography_field_runs():
                p_elem.append(bf_run)

            _save()
            return ok(f"Inserted bibliography field with heading '{heading_text}' at the end of the document. Press Ctrl+A then F9 in Word to populate.")

        else:
            return ok(f"Unknown tool: {name}")

    except Exception as exc:
        import traceback
        err_msg = str(exc)
        tb = traceback.format_exc()
        # Fallback suggestions for common errors
        suggestions = []
        if "No document open" in err_msg:
            suggestions.append("Open a document first: open_document(path='...') or create_new_document(path='...')")
        elif "index" in err_msg.lower() and ("range" in err_msg.lower() or "out of" in err_msg.lower()):
            suggestions.append("Use get_document_info or read_document to verify valid indices before editing.")
        elif "style" in err_msg.lower() and ("not found" in err_msg.lower() or "keyerror" in err_msg.lower()):
            suggestions.append("Use list_styles to see available styles, or create_style to make a new one.")
        elif "file not found" in err_msg.lower() or "no such file" in err_msg.lower():
            suggestions.append("Verify the file path exists and is a valid .docx file.")
        elif "permission" in err_msg.lower():
            suggestions.append("Ensure the file is not open in another program (e.g. Microsoft Word).")
        elif "msoffcrypto" in err_msg.lower():
            suggestions.append("Install missing dependency: pip install msoffcrypto-tool")
        elif "docx2pdf" in err_msg.lower():
            suggestions.append("Install missing dependency: pip install docx2pdf (Windows) or LibreOffice (Linux/macOS)")
        elif "lxml" in err_msg.lower():
            suggestions.append("Install missing dependency: pip install lxml")

        result = f"ERROR: {exc}\n{tb}"
        if suggestions:
            result += "\n\nFallback suggestions:\n" + '\n'.join(f"  - {s}" for s in suggestions)
        return ok(result)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream, write_stream,
            server.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())
