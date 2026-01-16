# OfficeMCP Enhancement Roadmap

## Overview
This document outlines the specifications for enhancing the OfficeMCP server with three major initiatives:
1. **Tool Consolidation & MCP Optimization** - Reduce tool explosion, leverage latest MCP features
2. **Sensitivity Label & Encryption Support** - Handle Microsoft Purview Information Protection labels and encrypted Office documents
3. **PDF Support** - Read and create PDF documents alongside Office documents

---

## MCP Server Best Practices & Latest Features

### Latest MCP Capabilities (Claude 3.5+ Compatible)

The Model Context Protocol has evolved with several important features that should be leveraged in this project:

#### 1. Tool Metadata & Annotations
Modern MCP servers support rich metadata to help models discover and use tools effectively:
- **Categories**: Group related tools (Core, Common, Advanced, Legacy)
- **Tags**: Label tools with capabilities (format-agnostic, destructive, read-only, sensitive, batch-capable)
- **Sampling Weights**: Control tool exposure frequency
- **Examples**: Provide real usage patterns inline
- **Related Tools**: Suggest complementary tools

```csharp
[McpServerTool(
    Name = "office_create",
    Category = ToolCategory.Core,
    Priority = 1,
    SamplingWeight = 1.0,
    Destructive = false
)]
[Tags(ToolTag.FormatAgnostic, ToolTag.Destructive)]
[Description("Creates a new document in any supported format")]
[Examples(new[] {
    "office_create(filePath='report.docx', title='Q4')",
    "office_create(filePath='data.xlsx')"
})]
[RelatedTools(new[] { "office_read", "office_write", "office_convert" })]
public string CreateDocument(...)
```

**Benefit**: Models can better understand tool purpose and make smarter choices

#### 2. Progressive Tool Disclosure
Instead of exposing all tools at once, use sampling configuration to progressively reveal tools:
- **Tier 1 (Always)**: Essential operations (5 tools)
- **Tier 2 (Contextual)**: Common operations (5 tools, shown when appropriate)
- **Tier 3 (On-Demand)**: Advanced operations (3 tools, shown on explicit request)

**Implementation**:
```json
{
  "tools": [
    { "name": "office_create", "samplingWeight": 1.0 },
    { "name": "office_read", "samplingWeight": 1.0 },
    { "name": "office_write", "samplingWeight": 1.0 },
    { "name": "office_convert", "samplingWeight": 1.0 },
    { "name": "office_metadata", "samplingWeight": 1.0 },
    { "name": "office_add_element", "samplingWeight": 0.7 },
    { "name": "office_add_header_footer", "samplingWeight": 0.6 },
    { "name": "office_extract", "samplingWeight": 0.7 },
    { "name": "office_encrypt", "samplingWeight": 0.4 },
    { "name": "office_batch", "samplingWeight": 0.5 }
  ]
}
```

**Benefit**: Reduces cognitive load; models focus on core tools first

#### 3. Input Validation & Constraints
Use JSON Schema to enforce parameter constraints at the protocol level:
```csharp
[McpServerTool]
public string ConvertDocument(
    [Description("Source file path")]
    [JsonPropertyName("sourcePath")]
    [StringLength(260)]  // Windows MAX_PATH
    [Pattern(@"^[a-zA-Z]:[\w\s\-./]*\.(docx|xlsx|pptx|pdf|md)$")]
    string sourcePath,
    
    [Description("Target format")]
    [EnumValues("docx", "xlsx", "pptx", "pdf", "md")]
    string targetFormat,
    
    [Description("Output path (optional)")]
    [StringLength(260)]
    string? outputPath = null
)
```

**Benefit**: Models get validation errors before execution; fewer errors

#### 4. Tool Grouping & Relationships
Define logical relationships between tools:
```csharp
public record ToolGroup(
    string Name,                          // "Document Reading"
    string Description,
    string[] ToolNames,                   // ["office_read", "office_extract"]
    string? PreferredEntryPoint = null    // "office_read" - start here
);
```

**Benefit**: Models understand workflows better

#### 5. Error Context & Recovery
Enhanced error responses help models recover:
```csharp
public record ErrorResult(
    bool Success,
    string ErrorMessage,
    string? ErrorCode = null,             // e.g., "FILE_NOT_FOUND"
    string? Suggestion = null,            // "Use office_create first"
    string[]? SuggestedTools = null,      // ["office_create"]
    object? Context = null                // File status, permissions, etc.
);
```

**Benefit**: Models can self-correct without human intervention

---

## Tool Consolidation & MCP Optimization

### Executive Summary
The current MCP server architecture exposes tools in a format-specific manner:
- `word_*` tools for Word documents
- `excel_*` tools for Excel documents  
- `ppt_*` tools for PowerPoint documents
- Proposed `pdf_*` tools for PDF documents

This approach creates tool explosion (30+ tools) that:
- **Increases LLM cognitive load** - Model must choose among many similar tools
- **Reduces discoverability** - Similar operations scattered across namespaces
- **Complicates maintenance** - Changes must be replicated across format implementations
- **Limits flexibility** - Format-specific logic prevents unified operations

### Proposed Solution: Format-Agnostic Unified Tools

#### Core Design Principles

1. **Single Tool Entry Points**
   - One tool per operation type (create, read, write, convert)
   - Tools auto-detect file format from extension or content
   - Format-specific behavior handled internally

2. **Format Agnostic Parameters**
   - Use `format` parameter to specify target format when needed
   - Supported formats: `docx`, `xlsx`, `pptx`, `pdf`, `md`
   - Auto-detection when reading (format inferred from file)

3. **Consistent Return Models**
   - All tools return format-agnostic results
   - Include format metadata in response
   - Enable easy chaining of operations

4. **Progressive Disclosure**
   - **Tier 1 (Core)**: Most common operations (5 tools)
   - **Tier 2 (Common)**: Format-specific enhancements (5 tools)
   - **Tier 3 (Advanced)**: Complex operations & batch (3 tools)
   - Total: ~13 tools vs. 30+ in current design

#### Unified Tool API

**Tier 1: Core Operations** (always exposed)
```
office_create:
  - filePath: string (required, format inferred from extension)
  - format: "docx|xlsx|pptx|pdf" (optional, inferred from path)
  - title: string? (optional, for document title)
  - markdown: string? (optional, initial content)
  - baseImagePath: string? (optional)
  - pageSize: "Letter|Legal|A4|A3" (default: Letter)
  ? Returns: DocumentResult

office_read:
  - filePath: string (required)
  - readType: "all|element|range" (default: all)
  - elementIndex: int? (for element mode: paragraph, slide, sheet)
  - startIndex/endIndex: int? (for range mode)
  - extractFormatting: bool (default: false, for PDF layout preservation)
  ? Returns: ContentResult with format metadata

office_write:
  - filePath: string (required, format inferred from extension)
  - content: string (required, markdown format)
  - targetFormat: "docx|xlsx|pptx|pdf" (optional, for format conversion)
  - baseImagePath: string? (optional)
  ? Returns: DocumentResult

office_convert:
  - sourcePath: string (required, any supported format)
  - targetFormat: "docx|xlsx|pptx|pdf|md" (required)
  - outputPath: string? (optional, auto-generated if not provided)
  ? Returns: DocumentResult with conversion metrics

office_metadata:
  - filePath: string (required)
  - includeStructure: bool (default: false, for detailed element listing)
  ? Returns: MetadataResult (unified across all formats)
```

**Tier 2: Common Operations** (enhanced capabilities)
```
office_add_element:
  - filePath: string (required)
  - elementType: "paragraph|heading|image|table|pageBreak|shape|chart|image|textbox" (required)
  - content: string? (required for text elements)
  - imagePath: string? (required for image elements)
  - tableData: string[][]? (required for table elements)
  - level: int? (for heading elements, 1-6)
  - formatting: object? (TextFormatting/TableFormatting as JSON)
  ? Returns: DocumentResult

office_add_header_footer:
  - filePath: string (required, Word/Excel/PDF only)
  - location: "header|footer" (required)
  - leftContent: string?
  - centerContent: string?
  - rightContent: string?
  - includePageNumber: bool (default: false)
  - includeDate: bool (default: false)
  ? Returns: DocumentResult

office_extract:
  - filePath: string (required)
  - extractType: "text|images|tables|metadata" (required)
  - scope: "all|element|range" (default: all)
  - elementIndex/startIndex/endIndex: int?
  ? Returns: ExtractionResult (polymorphic based on type)

office_encrypt:
  - filePath: string (required)
  - encryptionType: "password|sensitivity_label" (required)
  - userPassword: string? (for password encryption)
  - ownerPassword: string? (for owner-level password)
  - sensitivityLabel: string? (for label encryption, if supported)
  ? Returns: DocumentResult

office_decrypt:
  - filePath: string (required)
  - password: string? (for password-protected files)
  ? Returns: DocumentResult (with protection metadata)
```

**Tier 3: Advanced Operations** (specialized workflows)
```
office_batch:
  - filePath: string (required)
  - operations: object[] (required, JSON array of operations)
    Each operation can specify: type (markdown|element|convert|etc), format-specific params
  ? Returns: BatchOperationResult

office_merge:
  - outputPath: string (required, format inferred from extension)
  - inputPaths: string[] (required, JSON array)
  - preserveFormatting: bool (default: true)
  ? Returns: DocumentResult

office_template:
  - templatePath: string (required, source template)
  - outputPath: string (required, target document)
  - replacements: object (key-value pairs for template variables)
  ? Returns: DocumentResult

```
