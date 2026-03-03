# OfficeMCP

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server that enables AI assistants to create, read, write, and manipulate Microsoft Office documents (Word, Excel, PowerPoint) and PDFs directly from natural language instructions.

## Features

- **Format-agnostic tools** — unified `office_*` API works across all supported formats
- **Format-specific tools** — granular `word_*`, `excel_*`, and `pptx_*` tools for advanced operations
- **Markdown support** — write Word/PDF content using plain Markdown
- **Batch operations** — perform multiple document operations in a single tool call
- **PDF support** — create, read, watermark, and extract pages from PDFs
- **Encrypted document handling** — detection and support for protected Office files
- **Progressive tool disclosure** — tiered tool exposure keeps context lean
- **Image creation** — embed images in Word, PowerPoint, and Excel documents
- **Image extraction** — extract all embedded images from Word, PowerPoint, and PDF files as base64 data with MIME type for direct AI vision analysis (OCR, captioning)
- **Rich document reading** — `office_read` on `.docx` returns an ordered content list of headings, paragraphs, tables, and inline images in document sequence, preserving section context around every image

## Supported Formats

| Format | Extension | Read | Write | Create | Convert |
|--------|-----------|------|-------|--------|---------|
| Word | `.docx` | ✅ | ✅ | ✅ | ✅ |
| Excel | `.xlsx` | ✅ | ✅ | ✅ | — |
| PowerPoint | `.pptx` | ✅ | ✅ | ✅ | — |
| PDF | `.pdf` | ✅ | ✅ | ✅ | ✅ |
| Markdown | `.md` | — | — | — | ✅ |

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- An MCP-compatible client (e.g., Claude Desktop, VS Code with Copilot)

## Building

```bash
git clone https://github.com/your-org/OfficeMCP.git
cd OfficeMCP
dotnet build
```

## Running

```bash
dotnet run --project OfficeMCP/OfficeMCP.csproj
```

The server communicates over **stdio** using the MCP protocol.

## MCP Client Configuration

### Claude Desktop

Add the following to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "OfficeMCP": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "C:/path/to/OfficeMCP/OfficeMCP/OfficeMCP.csproj"
      ]
    }
  }
}
```

Or point to the compiled binary directly for faster startup:

```json
{
  "mcpServers": {
    "OfficeMCP": {
      "command": "C:/path/to/OfficeMCP/OfficeMCP/bin/Release/net10.0/OfficeMCP.exe"
    }
  }
}
```

## Available Tools

### Consolidated Office Tools (format-agnostic)

These tools work across all supported formats and are organized into three tiers.

#### Tier 1 — Core (always exposed)

| Tool | Description |
|------|-------------|
| `office_create` | Create a new document in any supported format |
| `office_read` | Read content from any document. For Word: returns an ordered content list with headings, paragraphs, tables, and inline images (base64) in reading order by default. Set `includeImages=false` for plain text only. |
| `office_write` | Write or append content to a document |
| `office_convert` | Convert between formats (e.g., `.docx` → `.md`) |
| `office_metadata` | Retrieve document metadata (format, size, structure) |

#### Tier 2 — Common

| Tool | Description |
|------|-------------|
| `office_add_element` | Add elements such as images, tables, and lists |
| `office_add_header_footer` | Add headers and footers with optional page numbers and dates |
| `office_extract` | Extract specific content: text, **images** (with base64 + context for AI OCR/captioning), tables, metadata |
| `office_batch` | Execute multiple operations on a document in one call |

#### Tier 3 — Advanced

| Tool | Description |
|------|-------------|
| `office_merge` | Merge multiple documents into one (PDF) |
| `office_pdf_pages` | Extract pages, add watermarks, or read specific PDF pages |

---

### Word Tools (`word_*`)

| Tool | Description |
|------|-------------|
| `word_read` | Read text from a Word document |
| `word_add_content` | Append Markdown content to a Word document |
| `word_add_element` | Add page breaks, headers, or footers |
| `word_add_image` | Embed an image (supports JPEG, PNG, GIF, BMP, TIFF) |
| `word_convert` | Convert between Word and Markdown |
| `word_batch` | Perform multiple write operations in one call |

---

### Excel Tools (`excel_*`)

| Tool | Description |
|------|-------------|
| `excel_create` | Create a workbook with optional initial table data |
| `excel_read` | Read content from a sheet, cell, or range |
| `excel_get_formatting` | Retrieve cell formatting (colors, fonts, borders) |
| `excel_set_cells` | Set cell values with optional table formatting |
| `excel_formula` | Insert a formula into a cell |
| `excel_manage_sheet` | Add, delete, or rename sheets |
| `excel_format_cells` | Apply formatting to a cell range |
| `excel_batch` | Perform multiple operations in one call |

---

### PowerPoint Tools (`pptx_*`)

| Tool | Description |
|------|-------------|
| `pptx_read` | Read text from slides |
| `pptx_add_slide` | Add a new blank slide |
| `pptx_manage_slide` | Delete, duplicate, or reorder slides |
| `pptx_add_title` | Set a slide title and subtitle |
| `pptx_add_text` | Add text or bullet points to a slide |
| `pptx_add_image` | Embed an image on a slide (supports JPEG, PNG, GIF, BMP) |
| `pptx_add_table` | Add a table to a slide |
| `pptx_add_notes` | Add speaker notes to a slide |
| `pptx_batch` | Perform multiple slide operations in one call |

## Image Support

### Creating documents with images

Images can be embedded when creating or editing documents:

| Format | How |
|--------|-----|
| Word | `word_add_image` tool, or inline Markdown `![alt](path)` via `office_create` / `office_write` |
| PowerPoint | `pptx_add_image` tool |
| Excel | `office_add_element` with `elementType: "image"` |

### Reading documents with images

#### Rich ordered reading (Word — default)

Calling `office_read` on a `.docx` file returns the full document as an **ordered content list** where each item has a `Type` of `"heading"`, `"paragraph"`, `"table"`, or `"image"`. Images appear immediately after their containing paragraph so the preceding headings and paragraphs provide natural section context — no separate extraction step needed.

```json
{
  "Content": [
    { "Type": "heading",   "Text": "Q4 Financial Summary", "Level": 1 },
    { "Type": "paragraph", "Text": "Revenue grew 23% as shown in the chart below." },
    { "Type": "image",     "AltText": "Revenue chart", "MimeType": "image/png",
                           "ImageBase64": "...", "WidthPx": 640, "HeightPx": 400 },
    { "Type": "paragraph", "Text": "Figure 1: Q4 2025 revenue by region." },
    { "Type": "heading",   "Text": "Cost Analysis", "Level": 2 }
  ]
}
```

Pass `includeImages: false` to get plain text only.

#### Bulk image extraction

`office_extract` with `extractType: "images"` extracts all embedded images from Word, PowerPoint, and PDF files. Each result includes:

| Field | Description |
|-------|-------------|
| `ImageBase64` | Full base64-encoded image bytes for AI vision analysis |
| `MimeType` | `image/png`, `image/jpeg`, etc. |
| `AltText` | Alt text stored in the document |
| `ContextBefore` | Paragraph immediately before the image (Word) or full slide text (PowerPoint) |
| `ContextAfter` | Paragraph immediately after the image (Word) |
| `WidthPx` / `HeightPx` | Dimensions at 96 dpi |
| `PageOrSlideNumber` | Source page (PDF) or slide number (PowerPoint) |

---

## Project Structure

```
OfficeMCP/
├── Program.cs                          # MCP server entry point & DI setup
├── Models/
│   └── DocumentModels.cs               # Shared data models (formatting, layout, etc.)
├── Services/
│   ├── WordDocumentService.cs          # Word document logic
│   ├── ExcelDocumentService.cs         # Excel workbook logic
│   ├── PowerPointDocumentService.cs    # PowerPoint presentation logic
│   ├── PdfDocumentService.cs           # PDF creation and manipulation
│   ├── EncryptedDocumentService.cs     # Encrypted/protected document handling
│   ├── FormatDetector.cs               # File extension → format detection
│   └── MarkdownParser.cs               # Markdown → Office content conversion
└── Tools/
    ├── OfficeDocumentToolsConsolidated.cs  # Unified office_* MCP tools
    ├── WordDocumentToolsOptimized.cs       # Word-specific word_* MCP tools
    ├── ExcelDocumentToolsOptimized.cs      # Excel-specific excel_* MCP tools
    └── PowerPointDocumentToolsOptimized.cs # PowerPoint-specific pptx_* MCP tools
```

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `ModelContextProtocol` | 1.0.0 | MCP server framework |
| `DocumentFormat.OpenXml` | 3.4.1 | Word, Excel, PowerPoint manipulation |
| `itext7` | 9.5.0 | PDF creation and manipulation |
| `Azure.Identity` | 1.18.0 | Azure credential support |
| `Microsoft.Extensions.Hosting` | 10.0.3 | Dependency injection and hosting |
| `Newtonsoft.Json` | 13.0.4 | JSON serialization (transitive override) |

## Running Tests

```bash
dotnet test OfficeMCP.Tests/OfficeMCP.Tests.csproj
```

## License

This project is provided as-is. 
