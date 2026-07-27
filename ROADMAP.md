# Research MCP — Roadmap

> **Vision**: Unified MCP server for the entire research lifecycle — discover, read, verify, store, map, visualize, write, generate figures, check plagiarism, humanize AI text, and produce publication-ready documents.

---

## Current State (Phase 0 — Complete)
**Word MCP Server** — 92 tools for Word document manipulation including editing, formatting, tables, images, headers/footers, bibliography workflow, batch operations, PDF export, and document builder.

---

## Phase 1: Paper Discovery & Acquisition
**Goal**: Find papers from web, fetch metadata, download PDFs.

**Tools**: `search_papers`, `get_paper_details`, `get_paper_citations`, `get_paper_references`, `get_author_info`, `get_author_papers`, `get_paper_recommendations`, `batch_get_papers`, `download_arxiv_source`, `download_pdf`, `verify_doi`, `get_citation_formats`

**Study**:
- [Silung/scholar-search-mcp](https://github.com/Silung/scholar-search-mcp) — Semantic Scholar + arXiv, 9 tools, MIT
- [kota-may478/academic-mcp-server-copilot](https://github.com/kota-may478/academic-mcp-server-copilot) — adds Crossref, full-text, MIT
- [isprime-coder/academic-mcp-server](https://github.com/isprime-coder/academic-mcp-server) — DOI verification, BibTeX, MIT

**APIs**: Semantic Scholar (free), arXiv (free), Crossref (free), OpenAlex (free), Unpaywall (free OA PDFs)

**Deps**: `httpx`, `feedparser`

---

## Phase 2: PDF Reading & Full-Text Extraction
**Goal**: Read PDFs, extract text, figures, tables.

**Tools**: `read_pdf`, `read_pdf_pages`, `get_pdf_outline`, `extract_figures`, `extract_tables`, `get_pdf_metadata`, `search_pdf_text`

**Study**:
- [zaddyzad/zotero-mcp](https://github.com/zaddyzad/zotero-mcp) — pymupdf extraction, MIT
- [Xevos117/mcp-zotero](https://github.com/Xevos117/mcp-zotero) — fulltext indexing, MIT

**Deps**: `pymupdf`, `pdfplumber`

---

## Phase 3: Reference Management & Storage (Zotero-like)
**Goal**: Store, organize, tag, and manage papers locally.

**Tools**: `add_to_library`, `search_library`, `get_library_item`, `list_collections`, `create_collection`, `add_to_collection`, `tag_item`, `find_duplicates`, `merge_duplicates`, `export_bibliography`, `import_bibliography`, `attach_pdf`, `get_item_fulltext`, `add_note`, `list_notes`, `update_note`, `delete_note`

**Study**:
- [Xevos117/mcp-zotero](https://github.com/Xevos117/mcp-zotero) — 15 tools, full Zotero API, MIT
- [zotero-library-mcp](https://pypi.org/project/zotero-library-mcp/) — DOI/arXiv/ISBN adding, BibTeX, MIT
- [iflow-mcp/x-t-e-r-zotero-mcp-neo](https://github.com/iflow-mcp/x-t-e-r-zotero-mcp-neo) — 6 unified tools, dryRun safety, MIT

**Architecture**: Local SQLite primary, optional Zotero sync. **Deps**: `sqlite3` (built-in), `pyzotero` (optional)

---

## Phase 4: Knowledge Graph & Visualization (Mind Palace)
**Goal**: Map papers, concepts, relationships into navigable graph. Visualize citation networks, concept maps.

**Tools**: `build_knowledge_graph`, `get_graph_subgraph`, `search_graph`, `add_concept`, `add_relationship`, `visualize_graph`, `visualize_mindmap`, `visualize_citation_network`, `visualize_concept_map`, `export_graph`

**Study**:
- [Sathvik-1007/GraphMem-MCP](https://github.com/Sathvik-1007/graphrag-mcp) — 28 tools, SQLite graph, vector+graph search, MIT
- [xsp52Hz/cognigraph-mcp-server](https://github.com/xsp52Hz/cognigraph-mcp-server) — mind maps, mermaid graphs, AI knowledge graphs, MIT
- [zcsabbagh/knowledge-graph-mcp](https://glama.ai/mcp/servers/zcsabbagh/knowledge-graph-mcp) — SQLite graph, Mermaid viz, MIT

**Viz tech**: Cytoscape.js, D3.js, Mermaid, markmap, Plotly. **Deps**: `networkx`, `pyvis`

---

## Phase 5: Scientific Figures, Charts & Diagrams
**Goal**: Generate publication-quality figures, charts, LaTeX equations.

**Tools**: `create_line_plot`, `create_bar_chart`, `create_scatter_plot`, `create_heatmap`, `create_histogram`, `create_box_plot`, `create_pie_chart`, `create_3d_plot`, `create_dashboard`, `render_latex_equation`, `render_tikz_diagram`, `create_flowchart`, `create_system_diagram`, `create_sequence_diagram`

**Study**:
- [Narroog/plot-MCP](https://github.com/Narroog/plot-MCP) — Nature-style plots, PDF+PNG, MIT
- [newsbubbles/matplotlib_mcp](https://github.com/newsbubbles/matplotlib_mcp) — full matplotlib, typed, MIT
- [viz-mcp](https://pypi.org/project/viz-mcp/) — auto-chart detection, Plotly HTML, MIT
- [tofunori/mcp-image-scientific](https://github.com/tofunori/mcp-image-scientific) — scientific figures via Gemini, MIT
- [danielsimonjr/upmath-mcp](https://github.com/danielsimonjr/upmath-mcp) — LaTeX/TikZ rendering, MIT

**Deps**: `matplotlib`, `plotly`, `seaborn`, `pandas`, `numpy`, `kaleido`

---

## Phase 6: Research Writing & Content Generation
**Goal**: Generate structured research content for different paper types.

**Tools**: `generate_outline`, `generate_abstract`, `generate_introduction`, `generate_literature_review`, `generate_methodology`, `generate_results_section`, `generate_discussion`, `generate_conclusion`, `generate_research_proposal`, `format_citation`, `generate_bibtex`, `check_structure`

**Study**:
- [ghostiee-11/overleaf-mcp](https://github.com/ghostiee-11/overleaf-mcp) — 24 tools, LaTeX check/compile, MIT
- [aaronsb/texflow-mcp](https://github.com/aaronsb/texflow-mcp) — structured doc model, MIT

**Paper types**: Empirical, review, theoretical, case study, proposal, thesis chapter, conference, journal

---

## Phase 7: AI Detection, Humanization & Plagiarism
**Goal**: Check text for AI patterns, plagiarism, and humanize AI-written content.

**Tools**: `detect_ai_content`, `get_ai_probability`, `analyze_writing_style`, `check_plagiarism`, `humanize_text`, `humanize_and_verify`, `compare_versions`, `batch_detect`, `get_detection_report`, `list_writing_styles`, `apply_style`

**Study**:
- [kitfoxs/humanize-mcp](https://github.com/kitfoxs/humanize-mcp) — 9-pass pipeline, 11 styles, local, MIT. **Best reference.**
- [zhaohongyuziranerran/ai-content-detector-mcp](https://glama.ai/mcp/servers/zhaohongyuziranerran/ai-content-detector-mcp) — 8 tools, detect+humanize loop, MIT

**Architecture**: Local-first (perplexity, burstiness, N-gram, pattern detection). Optional API (WriteHuman) for advanced.

---

## Phase 8: Citation Verification & Integrity
**Goal**: Verify citations exist, match claims, detect misrepresentation.

**Tools**: `verify_citation_exists`, `verify_citation_claim`, `find_uncited_sources`, `find_missing_citations`, `check_citation_consistency`, `verify_doi_batch`, `check_retracted_papers`, `get_citation_context`

**Study**: Scite API (citation reports, retractions), Retraction Watch API

---

## Phase 9: Workflow Orchestration
**Goal**: End-to-end research workflows tying all phases together.

**Workflows**: `research_workflow_search`, `research_workflow_review`, `research_workflow_write`, `research_workflow_verify`

Leverages existing 92 Word MCP tools for document output.

---

## Phase 10: Semantic Indexing
**Goal**: Local semantic search index over all stored papers.

**Tools**: `build_search_index`, `semantic_search`, `get_search_status`, `update_search_index`, `find_related_papers`, `cluster_papers`, `get_topic_summary`

**Study**: GraphMem-MCP (vector+graph RRF), jingkaimori/zotero-mcp (semantic search). **Deps**: `sentence-transformers` (optional), `scikit-learn`

---

## Build Order

| Priority | Phase | Effort |
|----------|-------|--------|
| 1 | Phase 1: Paper Discovery | Medium |
| 2 | Phase 2: PDF Reading | Medium |
| 3 | Phase 5: Figures & Charts | Medium (independent) |
| 4 | Phase 3: Reference Storage | Large |
| 5 | Phase 7: AI Detection | Medium (independent) |
| 6 | Phase 4: Knowledge Graph | Large |
| 7 | Phase 6: Writing & Generation | Large |
| 8 | Phase 8: Citation Verification | Medium |
| 9 | Phase 10: Semantic Indexing | Medium |
| 10 | Phase 9: Workflow Orchestration | Large |

---

## Architecture Principles

1. **Self-contained first** — local SQLite, local embeddings, local detection. No mandatory external services.
2. **Optional cloud** — Zotero sync, WriteHuman API, Semantic Scholar API key for higher limits.
3. **MIT licensed** — all reference projects are MIT. Maintain MIT license.
4. **Single server** — one `server.py`, one process, one MCP endpoint. All tools in one place.
5. **Incremental** — each phase is independently useful. No phase requires all prior phases.
6. **Credits** — acknowledge all upstream projects in README + LICENSE.

---

## License Summary (for going public)

All 3 current upstream repos use **MIT License** — permits merge, modify, publish, distribute with credit.

Reference projects for future phases (all MIT unless noted):
- scholar-search-mcp: MIT
- academic-mcp-server-copilot: MIT
- academic-mcp-server: MIT
- mcp-zotero (Xevos117): MIT
- zotero-library-mcp: MIT
- zotero-mcp-neo: MIT
- GraphMem-MCP: MIT
- cognigraph-mcp-server: MIT
- knowledge-graph-mcp: MIT
- plot-MCP: MIT
- matplotlib_mcp: MIT
- viz-mcp: MIT
- mcp-image-scientific: MIT
- upmath-mcp: MIT
- overleaf-mcp: MIT
- texflow-mcp: MIT
- humanize-mcp: MIT
- ai-content-detector-mcp: MIT
- Nexarag: CC BY 4.0 (paper), check code repo license

**Conclusion**: You can make the project public with credits. All core references are MIT.
