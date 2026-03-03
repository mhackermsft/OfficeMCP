using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// Unified MCP tools for all document formats (Word, Excel, PowerPoint, PDF).
/// Implements progressive tool disclosure with 3 tiers:
/// - Tier 1 (Core): Essential operations - always exposed
/// - Tier 2 (Common): Enhanced capabilities - shown when appropriate  
/// - Tier 3 (Advanced): Specialized workflows - shown on explicit request
/// </summary>
[McpServerToolType]
public sealed class OfficeDocumentToolsConsolidated(
    IServiceProvider serviceProvider,
    IEncryptedDocumentHandler? encryptionHandler = null)
{
    // Store for future encryption support (sensitivity labels, password protection)
    private readonly IEncryptedDocumentHandler? _encryptionHandler = encryptionHandler;
    
    private static readonly JsonSerializerOptions JsonOptions = 
        new() { PropertyNameCaseInsensitive = true, WriteIndented = false };

    #region TIER 1: CORE OPERATIONS (5 tools - always exposed)

    [McpServerTool(Name = "office_create", Destructive = false, ReadOnly = false), 
     Description("Creates a new document in any supported format (.docx, .xlsx, .pptx, .pdf). Format is auto-detected from file extension. For Excel, creates workbook with optional initial data. For PowerPoint, creates presentation with optional title slide. For Word: if you previously read a document with office_read, always pass the TemplatePath from that response to the templatePath parameter here to preserve the original styles, fonts, theme, and formatting.")]
    public string CreateDocument(
        [Description("Full path with extension (e.g., C:/docs/report.docx, data.xlsx, slides.pptx, document.pdf)")] 
        string filePath,
        [Description("Document title, sheet name, or slide title")] 
        string? title = null,
        [Description("Markdown content for Word/PDF initial content, or JSON 2D array for Excel initial data: [[\"Header1\",\"Header2\"],[\"Val1\",\"Val2\"]]")] 
        string? markdown = null,
        [Description("Base path for relative image paths in markdown")] 
        string? baseImagePath = null,
        [Description("Portrait or Landscape (Word/PDF only)")] 
        string orientation = "Portrait",
        [Description("Letter, Legal, A4, or A3 (Word/PDF only)")] 
        string pageSize = "Letter",
        [Description("Path to an existing .docx file to use as a style template. When provided, the new document inherits all styles, fonts, and theme from the source document, preserving the original formatting of headings, body text, etc.")] 
        string? templatePath = null)
    {
        try
        {
            var format = FormatDetector.DetectFormat(filePath);

            return format switch
            {
                "xlsx" => CreateExcelDocument(filePath, title, markdown),
                "pptx" => CreatePowerPointDocument(filePath, title),
                _ => CreateDocxOrPdfDocument(filePath, format, title, markdown, baseImagePath, orientation, pageSize, templatePath)
            };
        }
        catch (Exception ex)
        {
            return ErrorResult($"Creation failed: {ex.Message}", "Check file path is valid and directory exists");
        }
    }

    [McpServerTool(Name = "office_read", Destructive = false, ReadOnly = true),
     Description("Reads content from any document format. Returns all text by default. For Word (.docx): returns an ordered content list (headings, paragraphs, tables, inline images) and a TemplatePath field. When recreating or rewriting a Word document, always pass the TemplatePath value to office_create's templatePath parameter to preserve original styles, fonts, and theme. For Excel, reads all sheets or specific sheet/cell/range. For PowerPoint, reads all slides or specific slide.")]
    public string ReadDocument(
        [Description("Path to document (any supported format)")]
        string filePath,
        [Description("all (default), element/cell/slide, or range. For Excel: allSheets, sheet, cell, range")]
        string readType = "all",
        [Description("Element/paragraph/slide index (0-based), or sheet name for Excel")]
        string? elementId = null,
        [Description("Start index/cell (0-based for Word/PPT, cell reference like A1 for Excel)")]
        string? startRef = null,
        [Description("End index/cell (0-based for Word/PPT, cell reference like C10 for Excel)")]
        string? endRef = null,
        [Description("Word (.docx) only: when false, suppresses images and returns plain text only. Default is true: returns an ordered content list with headings, paragraphs, tables, and inline images (base64) in reading order so images appear in their section context.")]
        bool includeImages = true)
    {
        try
        {
            if (!File.Exists(filePath))
                return ErrorResult($"File not found: {filePath}", "Verify the file path is correct");

            var format = FormatDetector.DetectFormat(filePath);

            return format switch
            {
                "xlsx" => ReadExcelDocument(filePath, readType, elementId, startRef, endRef),
                "pptx" => ReadPowerPointDocument(filePath, readType, elementId),
                "docx" when includeImages => ReadDocxRichContent(filePath),
                _ => ReadDocxOrPdfDocument(filePath, format, readType, elementId, startRef, endRef)
            };
        }
        catch (Exception ex)
        {
            return ErrorResult($"Read failed: {ex.Message}");
        }
    }

    [McpServerTool(Name = "office_write", Destructive = false, ReadOnly = false),
     Description("Writes or appends content to a document. For Word/PDF: accepts markdown. For Excel: sets cell values. For PowerPoint: adds text to slides with optional positioning and formatting. For full PowerPoint control (shapes, lines, rich text, backgrounds), use the pptx_* tools directly.")]
    public string WriteDocument(
        [Description("Path to document")] 
        string filePath,
        [Description("Content: markdown for Word/PDF, cell value for Excel, text for PowerPoint")] 
        string content,
        [Description("For Excel: sheet name. For PowerPoint: slide index (0-based)")] 
        string? targetId = null,
        [Description("For Excel: cell reference (e.g., A1). For PowerPoint: X position in inches")] 
        string? position = null,
        [Description("Base path for relative image paths in markdown")] 
        string? baseImagePath = null,
        [Description("PowerPoint: Y position in inches (default 2.0)")] 
        double yInches = 2.0,
        [Description("PowerPoint: width in inches (default 8.0)")] 
        double widthInches = 8.0,
        [Description("PowerPoint: height in inches (default 1.0)")] 
        double heightInches = 1.0,
        [Description("PowerPoint: font size in points")] 
        int? fontSize = null,
        [Description("PowerPoint: font color as hex (e.g., 003366)")] 
        string? fontColor = null,
        [Description("PowerPoint: bold text")] 
        bool bold = false,
        [Description("PowerPoint: font name (e.g., Calibri, Segoe UI)")] 
        string? fontName = null)
    {
        try
        {
            if (!File.Exists(filePath))
                return ErrorResult($"File not found: {filePath}", "Use office_create to create the document first");

            if (string.IsNullOrWhiteSpace(content))
                return ErrorResult("Content cannot be empty");

            var format = FormatDetector.DetectFormat(filePath);

            return format switch
            {
                "xlsx" => WriteExcelDocument(filePath, content, targetId, position),
                "pptx" => WritePowerPointDocument(filePath, content, targetId, position, yInches, widthInches, heightInches, fontSize, fontColor, bold, fontName),
                _ => WriteDocxOrPdfDocument(filePath, format, content, baseImagePath)
            };
        }
        catch (Exception ex)
        {
            return ErrorResult($"Write failed: {ex.Message}");
        }
    }

    [McpServerTool(Name = "office_convert", Destructive = false, ReadOnly = false),
     Description("Converts document to markdown format. Full cross-format conversion coming soon.")]
    public string ConvertDocument(
        [Description("Path to source file")] 
        string sourcePath,
        [Description("Target format: md (markdown). Other formats coming soon.")] 
        string targetFormat,
        [Description("Output path (auto-generated if not provided)")] 
        string? outputPath = null)
    {
        try
        {
            if (!File.Exists(sourcePath))
                return ErrorResult($"Source file not found: {sourcePath}");

            if (!FormatDetector.IsSupported(targetFormat) && targetFormat.ToLowerInvariant() != "md")
                return ErrorResult($"Unsupported target format: {targetFormat}", "Supported: md (more coming soon)");

            var sourceFormat = FormatDetector.DetectFormat(sourcePath);

            if (targetFormat.ToLowerInvariant() == "md" && FormatDetector.UsesUnifiedInterface(sourceFormat))
            {
                var service = FormatDetector.GetService(sourceFormat, serviceProvider);
                var mdResult = service.ConvertToMarkdown(sourcePath);
                
                if (!mdResult.Success)
                    return JsonSerializer.Serialize(mdResult, JsonOptions);

                var mdPath = outputPath ?? Path.ChangeExtension(sourcePath, ".md");
                EnsureDirectoryExists(mdPath);
                File.WriteAllText(mdPath, mdResult.Content);
                
                return SuccessResult($"Converted to markdown", mdPath, sourceFormat);
            }

            return ErrorResult($"Conversion from {sourceFormat} to {targetFormat} not yet implemented",
                "Currently only markdown export is supported for Word and PDF");
        }
        catch (Exception ex)
        {
            return ErrorResult($"Conversion failed: {ex.Message}");
        }
    }

    [McpServerTool(Name = "office_metadata", Destructive = false, ReadOnly = true),
     Description("Gets metadata about a document including format, size, dates, and element counts.")]
    public string GetMetadata(
        [Description("Path to document")] 
        string filePath,
        [Description("Include detailed structure (element counts, sheet names, slide count)")] 
        bool includeStructure = false)
    {
        try
        {
            if (!File.Exists(filePath))
                return ErrorResult($"File not found: {filePath}");

            var format = FormatDetector.DetectFormat(filePath);
            var fileInfo = new FileInfo(filePath);

            var metadata = new Dictionary<string, object>
            {
                ["Success"] = true,
                ["Format"] = format,
                ["FilePath"] = filePath,
                ["FileSize"] = fileInfo.Length,
                ["FileSizeFormatted"] = FormatFileSize(fileInfo.Length),
                ["CreatedDate"] = fileInfo.CreationTime,
                ["ModifiedDate"] = fileInfo.LastWriteTime
            };

            if (includeStructure)
            {
                try
                {
                    switch (format)
                    {
                        case "xlsx":
                            var excelService = FormatDetector.GetExcelService(serviceProvider);
                            var sheetsResult = excelService.GetAllSheetsText(filePath);
                            metadata["SheetInfo"] = sheetsResult.Content ?? "Unable to read sheets";
                            break;
                        case "pptx":
                            var pptService = FormatDetector.GetPowerPointService(serviceProvider);
                            var countResult = pptService.GetSlideCount(filePath);
                            metadata["SlideCount"] = countResult.Content ?? "Unable to count slides";
                            break;
                        case "pdf":
                            var pdfService = FormatDetector.GetPdfService(serviceProvider);
                            var pdfResult = pdfService.GetDocumentText(filePath);
                            metadata["PageCount"] = pdfResult.TotalPages ?? 0;
                            break;
                        default:
                            var docService = FormatDetector.GetService(format, serviceProvider);
                            var docResult = docService.GetDocumentText(filePath);
                            metadata["ParagraphCount"] = docResult.TotalParagraphs ?? 0;
                            break;
                    }
                }
                catch { /* Structure info is optional */ }
            }

            return JsonSerializer.Serialize(metadata, JsonOptions);
        }
        catch (Exception ex)
        {
            return ErrorResult($"Metadata retrieval failed: {ex.Message}");
        }
    }

    #endregion

    #region TIER 2: COMMON OPERATIONS (5 tools - shown when appropriate)

    [McpServerTool(Name = "office_add_element", Destructive = false, ReadOnly = false),
     Description("Adds content elements to a document. Supports: paragraph, heading, image, table, pageBreak, bulletList, numberedList, shape, line. For PowerPoint: supports X/Y positioning, font styling, and colors. For full PowerPoint control (rich text, gradients, connectors, groups), use the pptx_* tools directly.")]
    public string AddElement(
        [Description("Path to document")] 
        string filePath,
        [Description("Element type: paragraph, heading, image, table, pageBreak, bulletList, numberedList. PowerPoint also: shape, line")] 
        string elementType,
        [Description("Text content (for text elements) or JSON table data: [[\"H1\",\"H2\"],[\"R1C1\",\"R1C2\"]]")] 
        string? content = null,
        [Description("Image file path (for image elements)")] 
        string? imagePath = null,
        [Description("Heading level 1-6, or slide index for PowerPoint")] 
        int level = 1,
        [Description("Width in inches (images/tables/shapes)")] 
        double widthInches = 4.0,
        [Description("Height in inches (images/shapes)")] 
        double heightInches = 3.0,
        [Description("Sheet name for Excel, ignored for other formats")] 
        string? sheetName = null,
        [Description("Start cell for Excel tables (e.g., A1)")] 
        string? startCell = null,
        [Description("PowerPoint: X position in inches (default 1.0)")] 
        double xInches = 1.0,
        [Description("PowerPoint: Y position in inches (default 2.0)")] 
        double yInches = 2.0,
        [Description("PowerPoint: font size in points")] 
        int? fontSize = null,
        [Description("PowerPoint: font color as hex (e.g., 003366)")] 
        string? fontColor = null,
        [Description("PowerPoint: fill/background color as hex")] 
        string? fillColor = null,
        [Description("PowerPoint: bold text")] 
        bool bold = false,
        [Description("PowerPoint: font name (e.g., Calibri)")] 
        string? fontName = null,
        [Description("PowerPoint: shape type (rectangle, roundRectangle, ellipse, etc.) for shape element")] 
        string? shapeType = null)
    {
        try
        {
            if (!File.Exists(filePath))
                return ErrorResult($"File not found: {filePath}");

            var format = FormatDetector.DetectFormat(filePath);

            // Handle table element type with JSON data
            if (elementType.Equals("table", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(content))
            {
                return AddTableElement(filePath, format, content, sheetName, startCell, level);
            }

            // Handle format-specific element addition
            return format switch
            {
                "xlsx" => AddExcelElement(filePath, elementType, content, imagePath, sheetName, startCell),
                "pptx" => AddPowerPointElement(filePath, elementType, content, imagePath, level, widthInches, heightInches,
                    xInches, yInches, fontSize, fontColor, fillColor, bold, fontName, shapeType),
                _ => AddDocumentElement(filePath, format, elementType, content, imagePath, level, widthInches, heightInches)
            };
        }
        catch (Exception ex)
        {
            return ErrorResult($"Add element failed: {ex.Message}");
        }
    }

    [McpServerTool(Name = "office_add_header_footer", Destructive = false, ReadOnly = false),
     Description("Adds header or footer to Word or PDF documents. Supports left/center/right content, page numbers, and dates.")]
    public string AddHeaderFooter(
        [Description("Path to Word or PDF document")] 
        string filePath,
        [Description("header or footer")] 
        string location,
        [Description("Left-aligned content")] 
        string? leftContent = null,
        [Description("Center-aligned content")] 
        string? centerContent = null,
        [Description("Right-aligned content")] 
        string? rightContent = null,
        [Description("Include automatic page number")] 
        bool includePageNumber = false,
        [Description("Include current date")] 
        bool includeDate = false)
    {
        try
        {
            if (!File.Exists(filePath))
                return ErrorResult($"File not found: {filePath}");

            var format = FormatDetector.DetectFormat(filePath);
            
            if (format != "docx" && format != "pdf")
                return ErrorResult($"Header/footer not supported for {format}", "Only Word (.docx) and PDF (.pdf) support headers/footers");

            var service = FormatDetector.GetService(format, serviceProvider);
            var options = new HeaderFooterOptions(leftContent, centerContent, rightContent, includePageNumber, includeDate);

            var result = location.ToLowerInvariant() switch
            {
                "header" => service.AddHeader(filePath, options),
                "footer" => service.AddFooter(filePath, options),
                _ => new DocumentResult(false, $"Invalid location: {location}", Suggestion: "Use 'header' or 'footer'")
            };

            return JsonSerializer.Serialize(result with { Format = format }, JsonOptions);
        }
        catch (Exception ex)
        {
            return ErrorResult($"Add header/footer failed: {ex.Message}");
        }
    }

    [McpServerTool(Name = "office_extract", Destructive = false, ReadOnly = true),
     Description("Extracts specific content types from documents: text, images, tables, or metadata.")]
    public string ExtractContent(
        [Description("Path to document")] 
        string filePath,
        [Description("What to extract: text, images, tables, metadata")] 
        string extractType,
        [Description("Scope: all (default), element, or range")] 
        string scope = "all",
        [Description("Element index for element scope, or start index for range")] 
        int? startIndex = null,
        [Description("End index for range scope")] 
        int? endIndex = null)
    {
        try
        {
            if (!File.Exists(filePath))
                return ErrorResult($"File not found: {filePath}");

            var format = FormatDetector.DetectFormat(filePath);

            return extractType.ToLowerInvariant() switch
            {
                "text" => ReadDocument(filePath, scope == "element" ? "element" : scope == "range" ? "range" : "all",
                    startIndex?.ToString(), startIndex?.ToString(), endIndex?.ToString()),
                "metadata" => GetMetadata(filePath, true),
                "images" => ExtractDocumentImages(filePath, format),
                "tables" => ErrorResult("Table extraction not yet implemented", "Coming in future release"),
                _ => ErrorResult($"Unknown extract type: {extractType}", "Supported: text, images, tables, metadata")
            };
        }
        catch (Exception ex)
        {
            return ErrorResult($"Extraction failed: {ex.Message}");
        }
    }

    #endregion

    #region TIER 3: ADVANCED OPERATIONS (3 tools - specialized workflows)

    [McpServerTool(Name = "office_batch", Destructive = false, ReadOnly = false),
     Description("Executes multiple operations on a document in a single call. More efficient than multiple individual calls.")]
    public string BatchOperations(
        [Description("Path to document")] 
        string filePath,
        [Description("JSON array of operations: [{\"type\":\"paragraph\",\"content\":\"text\"},{\"type\":\"heading\",\"content\":\"title\",\"level\":1}]")] 
        string operationsJson)
    {
        try
        {
            if (!File.Exists(filePath))
                return ErrorResult($"File not found: {filePath}", "Use office_create first");

            var operations = JsonSerializer.Deserialize<JsonElement[]>(operationsJson, JsonOptions);
            if (operations == null || operations.Length == 0)
                return ErrorResult("No operations provided", "Provide JSON array of operations");

            var format = FormatDetector.DetectFormat(filePath);
            var results = new List<OperationOutcome>();
            int successful = 0, failed = 0;

            for (int i = 0; i < operations.Length; i++)
            {
                var op = operations[i];
                var opType = op.TryGetProperty("type", out var typeEl) ? typeEl.GetString() ?? "" : "";
                
                try
                {
                    var opResult = ExecuteBatchOperation(filePath, format, op, opType);
                    var success = opResult.Contains("\"Success\":true", StringComparison.OrdinalIgnoreCase);
                    
                    if (success) successful++; else failed++;
                    results.Add(new OperationOutcome(i, opType, success, success ? "OK" : "Failed"));
                }
                catch (Exception opEx)
                {
                    failed++;
                    results.Add(new OperationOutcome(i, opType, false, opEx.Message));
                }
            }

            return JsonSerializer.Serialize(new BatchOperationResult(
                failed == 0, 
                failed == 0 ? "All operations completed" : $"{failed} operation(s) failed",
                operations.Length, successful, failed, results), JsonOptions);
        }
        catch (Exception ex)
        {
            return ErrorResult($"Batch operation failed: {ex.Message}", "Check JSON format is valid");
        }
    }

    [McpServerTool(Name = "office_merge", Destructive = false, ReadOnly = false),
     Description("Merges multiple documents into one. Currently supports PDF merging.")]
    public string MergeDocuments(
        [Description("Output file path (format detected from extension)")] 
        string outputPath,
        [Description("JSON array of input file paths: [\"file1.pdf\",\"file2.pdf\"]")] 
        string inputPathsJson,
        [Description("Preserve original formatting where possible")] 
        bool preserveFormatting = true)
    {
        try
        {
            var inputPaths = JsonSerializer.Deserialize<string[]>(inputPathsJson, JsonOptions);
            if (inputPaths == null || inputPaths.Length == 0)
                return ErrorResult("No input files provided");

            var format = FormatDetector.DetectFormat(outputPath);

            if (format == "pdf")
            {
                var pdfService = FormatDetector.GetPdfService(serviceProvider);
                var result = pdfService.MergeDocuments(outputPath, inputPaths);
                return JsonSerializer.Serialize(result with { Format = format }, JsonOptions);
            }

            return ErrorResult($"Merge not supported for {format}", "Currently only PDF merging is supported");
        }
        catch (Exception ex)
        {
            return ErrorResult($"Merge failed: {ex.Message}");
        }
    }

    [McpServerTool(Name = "office_pdf_pages", Destructive = false, ReadOnly = false),
     Description("PDF-specific operations: extract specific pages, add watermark, get page text.")]
    public string PdfPageOperations(
        [Description("Path to PDF file")] 
        string filePath,
        [Description("Operation: extract_pages, watermark, get_page")] 
        string operation,
        [Description("For extract_pages: JSON array of page numbers [1,3,5]. For get_page: page number")] 
        string? pageNumbers = null,
        [Description("Output path for extract_pages")] 
        string? outputPath = null,
        [Description("Watermark text")] 
        string? watermarkText = null,
        [Description("Watermark opacity 0.0-1.0")] 
        double opacity = 0.3,
        [Description("Watermark rotation in degrees")] 
        double rotation = -45)
    {
        try
        {
            if (!File.Exists(filePath))
                return ErrorResult($"File not found: {filePath}");

            var format = FormatDetector.DetectFormat(filePath);
            if (format != "pdf")
                return ErrorResult("This tool only works with PDF files", $"File is {format}, not PDF");

            var pdfService = FormatDetector.GetPdfService(serviceProvider);

            return operation.ToLowerInvariant() switch
            {
                "extract_pages" when !string.IsNullOrEmpty(pageNumbers) && !string.IsNullOrEmpty(outputPath) =>
                    ExtractPdfPages(pdfService, filePath, pageNumbers, outputPath),
                "extract_pages" =>
                    ErrorResult("pageNumbers and outputPath required for extract_pages"),
                
                "watermark" when !string.IsNullOrEmpty(watermarkText) =>
                    JsonSerializer.Serialize(pdfService.AddWatermark(filePath, watermarkText, 
                        new WatermarkOptions(opacity, rotation)), JsonOptions),
                "watermark" =>
                    ErrorResult("watermarkText is required for watermark operation"),
                
                "get_page" when int.TryParse(pageNumbers, out var pageNum) =>
                    JsonSerializer.Serialize(pdfService.GetPageText(filePath, pageNum), JsonOptions),
                "get_page" =>
                    ErrorResult("Valid page number required for get_page"),
                
                _ => ErrorResult($"Unknown operation: {operation}", "Supported: extract_pages, watermark, get_page")
            };
        }
        catch (Exception ex)
        {
            return ErrorResult($"PDF operation failed: {ex.Message}");
        }
    }

    #endregion

    #region Private Helper Methods

    private string CreateExcelDocument(string filePath, string? sheetName, string? initialDataJson)
    {
        var excelService = FormatDetector.GetExcelService(serviceProvider);
        var result = excelService.CreateWorkbook(filePath, sheetName);
        
        if (!result.Success)
            return JsonSerializer.Serialize(result with { Format = "xlsx" }, JsonOptions);

        if (!string.IsNullOrWhiteSpace(initialDataJson))
        {
            try
            {
                var data = JsonSerializer.Deserialize<string[][]>(initialDataJson, JsonOptions);
                if (data != null && data.Length > 0)
                {
                    var tableResult = excelService.AddTable(filePath, sheetName ?? "Sheet1", "A1", data, true);
                    if (!tableResult.Success)
                        return ErrorResult($"Created but data failed: {tableResult.Message}", filePath);
                }
            }
            catch (JsonException)
            {
                // If not valid JSON, treat as markdown (not applicable for Excel)
            }
        }

        return SuccessResult("Workbook created", filePath, "xlsx");
    }

    private string CreatePowerPointDocument(string filePath, string? title)
    {
        var pptService = FormatDetector.GetPowerPointService(serviceProvider);
        var result = pptService.CreatePresentation(filePath, title);
        return JsonSerializer.Serialize(result with { Format = "pptx" }, JsonOptions);
    }

    private string CreateDocxOrPdfDocument(string filePath, string format, string? title, string? markdown,
        string? baseImagePath, string orientation, string pageSize, string? templatePath = null)
    {
        var service = FormatDetector.GetService(format, serviceProvider);
        var layout = new PageLayoutOptions(Orientation: orientation, PageSize: pageSize);
        var result = service.CreateDocument(filePath, title, layout, templatePath);

        if (!result.Success)
            return JsonSerializer.Serialize(result with { Format = format }, JsonOptions);

        if (!string.IsNullOrWhiteSpace(markdown))
        {
            var mdResult = service.AddMarkdownContent(filePath, markdown, baseImagePath);
            if (!mdResult.Success)
                return ErrorResult($"Created but content failed: {mdResult.Message}", filePath);
        }

        return SuccessResult("Document created", filePath, format);
    }

    private string ReadExcelDocument(string filePath, string readType, string? sheetName, string? startCell, string? endCell)
    {
        var excelService = FormatDetector.GetExcelService(serviceProvider);
        
        ContentResult result = readType.ToLowerInvariant() switch
        {
            "sheet" when !string.IsNullOrEmpty(sheetName) => excelService.GetSheetText(filePath, sheetName),
            "cell" when !string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(startCell) =>
                excelService.GetCellValue(filePath, sheetName, startCell),
            "range" when !string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(startCell) && !string.IsNullOrEmpty(endCell) =>
                excelService.GetRangeValues(filePath, sheetName, startCell, endCell),
            _ => excelService.GetAllSheetsText(filePath)
        };

        return JsonSerializer.Serialize(result with { Format = "xlsx" }, JsonOptions);
    }

    private string ReadPowerPointDocument(string filePath, string readType, string? slideIndex)
    {
        var pptService = FormatDetector.GetPowerPointService(serviceProvider);
        
        ContentResult result = readType.ToLowerInvariant() switch
        {
            "slide" or "element" when int.TryParse(slideIndex, out var idx) => pptService.GetSlideText(filePath, idx),
            "count" => pptService.GetSlideCount(filePath),
            _ => pptService.GetAllSlidesText(filePath)
        };

        return JsonSerializer.Serialize(result with { Format = "pptx" }, JsonOptions);
    }

    private string ExtractDocumentImages(string filePath, string format)
    {
        IList<OfficeMCP.Models.ImageExtractionResult> images = format switch
        {
            "pptx" => FormatDetector.GetPowerPointService(serviceProvider).ExtractImages(filePath),
            _ => FormatDetector.GetService(format, serviceProvider).ExtractImages(filePath)
        };

        var response = new
        {
            Success = true,
            Format = format,
            ImageCount = images.Count,
            Note = images.Count == 0
                ? "No embedded images found in this document"
                : "Each image includes base64-encoded data (ImageBase64), alt text, and surrounding text context. " +
                  "Use the image data for OCR and caption generation.",
            Images = images
        };
        return JsonSerializer.Serialize(response, JsonOptions);
    }

    private string ReadDocxRichContent(string filePath)
    {
        var wordService = (IWordDocumentService)FormatDetector.GetService("docx", serviceProvider);
        var items = wordService.GetRichContent(filePath);
        var response = new
        {
            Success = true,
            Format = "docx",
            FilePath = filePath,
            TemplatePath = filePath,
            ItemCount = items.Count,
            Note = "Content is in document reading order. Images appear immediately after their containing paragraph. " +
                   "Preceding headings and paragraphs provide section context. " +
                   "Use MimeType + ImageBase64 for AI vision analysis, OCR, and caption generation. " +
                   "IMPORTANT: When recreating or rewriting this document, pass the TemplatePath value to the templatePath parameter of office_create to preserve the original styles, fonts, and formatting.",
            Content = items
        };
        return JsonSerializer.Serialize(response, JsonOptions);
    }

    private string ReadDocxOrPdfDocument(string filePath, string format, string readType, string? elementId, string? startRef, string? endRef)
    {
        var service = FormatDetector.GetService(format, serviceProvider);
        
        ContentResult result = readType.ToLowerInvariant() switch
        {
            "element" when int.TryParse(elementId, out var idx) => service.GetParagraphText(filePath, idx),
            "range" when int.TryParse(startRef, out var start) && int.TryParse(endRef, out var end) =>
                service.GetParagraphRange(filePath, start, end),
            _ => service.GetDocumentText(filePath)
        };

        return JsonSerializer.Serialize(result with { Format = format }, JsonOptions);
    }

    private string WriteExcelDocument(string filePath, string value, string? sheetName, string? cellRef)
    {
        var excelService = FormatDetector.GetExcelService(serviceProvider);
        var sheet = sheetName ?? "Sheet1";
        var cell = cellRef ?? "A1";
        
        var result = excelService.SetCellValue(filePath, sheet, cell, value);
        return JsonSerializer.Serialize(result with { Format = "xlsx" }, JsonOptions);
    }

    private string WritePowerPointDocument(string filePath, string text, string? slideIndex, string? position,
        double yInches = 2.0, double widthInches = 8.0, double heightInches = 1.0,
        int? fontSize = null, string? fontColor = null, bool bold = false, string? fontName = null)
    {
        var pptService = FormatDetector.GetPowerPointService(serviceProvider);
        var idx = int.TryParse(slideIndex, out var i) ? i : 0;
        var x = double.TryParse(position, out var p) ? p : 1.0;
        
        TextFormatting? textFormat = (fontSize.HasValue || fontColor != null || bold || fontName != null)
            ? new TextFormatting(Bold: bold, FontSize: fontSize, FontColor: fontColor, FontName: fontName)
            : null;

        var options = new TextBoxOptions(
            X: (long)(x * 914400),
            Y: (long)(yInches * 914400),
            Width: (long)(widthInches * 914400),
            Height: (long)(heightInches * 914400),
            TextFormat: textFormat);
        
        var result = pptService.AddTextBox(filePath, idx, text, options);
        return JsonSerializer.Serialize(result with { Format = "pptx" }, JsonOptions);
    }

    private string WriteDocxOrPdfDocument(string filePath, string format, string markdown, string? baseImagePath)
    {
        var service = FormatDetector.GetService(format, serviceProvider);
        var result = service.AddMarkdownContent(filePath, markdown, baseImagePath);
        return JsonSerializer.Serialize(result with { FilePath = filePath, Format = format }, JsonOptions);
    }

    private string AddTableElement(string filePath, string format, string tableDataJson, string? sheetName, string? startCell, int slideIndex)
    {
        try
        {
            var data = JsonSerializer.Deserialize<string[][]>(tableDataJson, JsonOptions);
            if (data == null || data.Length == 0)
                return ErrorResult("Invalid table data");

            return format switch
            {
                "xlsx" => AddExcelTable(filePath, data, sheetName, startCell),
                "pptx" => AddPowerPointTable(filePath, data, slideIndex),
                _ => AddDocumentTable(filePath, format, data)
            };
        }
        catch (JsonException)
        {
            return ErrorResult("Invalid JSON table data", "Format: [[\"H1\",\"H2\"],[\"R1C1\",\"R1C2\"]]");
        }
    }

    private string AddExcelTable(string filePath, string[][] data, string? sheetName, string? startCell)
    {
        var excelService = FormatDetector.GetExcelService(serviceProvider);
        var result = excelService.AddTable(filePath, sheetName ?? "Sheet1", startCell ?? "A1", data, true);
        return JsonSerializer.Serialize(result with { Format = "xlsx" }, JsonOptions);
    }

    private string AddPowerPointTable(string filePath, string[][] data, int slideIndex)
    {
        var pptService = FormatDetector.GetPowerPointService(serviceProvider);
        var result = pptService.AddTable(filePath, slideIndex, data,
            (long)(1.0 * 914400), (long)(2.0 * 914400), (long)(8.0 * 914400), (long)(3.0 * 914400));
        return JsonSerializer.Serialize(result with { Format = "pptx" }, JsonOptions);
    }

    private string AddDocumentTable(string filePath, string format, string[][] data)
    {
        var service = FormatDetector.GetService(format, serviceProvider);
        var result = service.AddTable(filePath, data);
        return JsonSerializer.Serialize(result with { Format = format }, JsonOptions);
    }

    private string AddExcelElement(string filePath, string elementType, string? content, string? imagePath, string? sheetName, string? startCell)
    {
        var excelService = FormatDetector.GetExcelService(serviceProvider);
        var sheet = sheetName ?? "Sheet1";
        var cell = startCell ?? "A1";

        return elementType.ToLowerInvariant() switch
        {
            "image" when !string.IsNullOrEmpty(imagePath) && File.Exists(imagePath) =>
                JsonSerializer.Serialize(excelService.AddImage(filePath, sheet, imagePath, cell), JsonOptions),
            "image" => ErrorResult("Image path required and must exist"),
            _ => ErrorResult($"Element type '{elementType}' not supported for Excel", "Use 'image' or 'table'")
        };
    }

    private string AddPowerPointElement(string filePath, string elementType, string? content, string? imagePath, 
        int slideIndex, double width, double height,
        double xInches = 1.0, double yInches = 2.0, int? fontSize = null, string? fontColor = null,
        string? fillColor = null, bool bold = false, string? fontName = null, string? shapeType = null)
    {
        var pptService = FormatDetector.GetPowerPointService(serviceProvider);
        var xEmu = (long)(xInches * 914400);
        var yEmu = (long)(yInches * 914400);
        var wEmu = (long)(width * 914400);
        var hEmu = (long)(height * 914400);

        TextFormatting? textFormat = (fontSize.HasValue || fontColor != null || bold || fontName != null)
            ? new TextFormatting(Bold: bold, FontSize: fontSize, FontColor: fontColor, FontName: fontName)
            : null;

        return elementType.ToLowerInvariant() switch
        {
            "image" when !string.IsNullOrEmpty(imagePath) && File.Exists(imagePath) =>
                JsonSerializer.Serialize(pptService.AddImage(filePath, slideIndex, imagePath,
                    xEmu, yEmu, new ImageOptions(wEmu, hEmu)), JsonOptions),
            "image" => ErrorResult("Image path required and must exist"),
            "paragraph" or "text" when !string.IsNullOrEmpty(content) =>
                JsonSerializer.Serialize(pptService.AddTextBox(filePath, slideIndex, content,
                    new TextBoxOptions(xEmu, yEmu, wEmu, hEmu, TextFormat: textFormat, BackgroundColor: fillColor)), JsonOptions),
            "heading" or "title" when !string.IsNullOrEmpty(content) =>
                JsonSerializer.Serialize(pptService.AddTitle(filePath, slideIndex, content, null, 
                    textFormat ?? new TextFormatting(Bold: true, FontSize: 44)), JsonOptions),
            "bulletlist" when !string.IsNullOrEmpty(content) =>
                JsonSerializer.Serialize(pptService.AddBulletPoints(filePath, slideIndex, content.Split('\n'),
                    new TextBoxOptions(xEmu, yEmu, wEmu, (long)(4.0 * 914400), TextFormat: textFormat)), JsonOptions),
            "shape" when !string.IsNullOrEmpty(content) || !string.IsNullOrEmpty(shapeType) =>
                JsonSerializer.Serialize(pptService.AddShape(filePath, slideIndex,
                    new ShapeOptions(ShapeType: shapeType ?? "rectangle",
                        X: xEmu, Y: yEmu, Width: wEmu, Height: hEmu,
                        Text: content, FillColor: fillColor ?? "4472C4", 
                        TextFormat: textFormat)), JsonOptions),
            "line" =>
                JsonSerializer.Serialize(pptService.AddLine(filePath, slideIndex,
                    new LineOptions(X1: xEmu, Y1: yEmu,
                        X2: (long)((xInches + width) * 914400), Y2: yEmu)), JsonOptions),
            _ => ErrorResult($"Element type '{elementType}' requires content", "Provide 'content' parameter")
        };
    }

    private string AddDocumentElement(string filePath, string format, string elementType, string? content, string? imagePath, int level, double width, double height)
    {
        var service = FormatDetector.GetService(format, serviceProvider);

        DocumentResult result = elementType.ToLowerInvariant() switch
        {
            "paragraph" when !string.IsNullOrWhiteSpace(content) => service.AddParagraph(filePath, content),
            "heading" when !string.IsNullOrWhiteSpace(content) => service.AddHeading(filePath, content, level),
            "image" when !string.IsNullOrWhiteSpace(imagePath) && File.Exists(imagePath) =>
                service.AddImage(filePath, imagePath, new ImageOptions((long)(width * 914400), (long)(height * 914400))),
            "pagebreak" => service.AddPageBreak(filePath),
            "bulletlist" when !string.IsNullOrWhiteSpace(content) => service.AddBulletList(filePath, content.Split('\n')),
            "numberedlist" when !string.IsNullOrWhiteSpace(content) => service.AddNumberedList(filePath, content.Split('\n')),
            _ => new DocumentResult(false, $"Element type '{elementType}' requires content or valid image path")
        };

        return JsonSerializer.Serialize(result with { Format = format }, JsonOptions);
    }

    private string ExecuteBatchOperation(string filePath, string format, JsonElement op, string opType)
    {
        var content = op.TryGetProperty("content", out var c) ? c.GetString() : null;
        var text = op.TryGetProperty("text", out var t) ? t.GetString() : content;
        var level = op.TryGetProperty("level", out var l) ? l.GetInt32() : 1;
        var imagePath = op.TryGetProperty("imagePath", out var img) ? img.GetString() : null;

        return opType.ToLowerInvariant() switch
        {
            "paragraph" or "heading" or "image" or "pagebreak" or "bulletlist" or "numberedlist" or "table" =>
                AddElement(filePath, opType, text, imagePath, level),
            "markdown" when !string.IsNullOrEmpty(text) =>
                WriteDocument(filePath, text),
            _ => ErrorResult($"Unknown batch operation type: {opType}")
        };
    }

    private string ExtractPdfPages(IPdfDocumentService pdfService, string filePath, string pageNumbersJson, string outputPath)
    {
        var pages = JsonSerializer.Deserialize<int[]>(pageNumbersJson, JsonOptions);
        if (pages == null || pages.Length == 0)
            return ErrorResult("Invalid page numbers");

        var result = pdfService.ExtractPages(filePath, pages, outputPath);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string ErrorResult(string message, string? suggestion = null)
    {
        return JsonSerializer.Serialize(new { Success = false, Message = message, Suggestion = suggestion }, JsonOptions);
    }

    private static string SuccessResult(string message, string? filePath = null, string? format = null)
    {
        return JsonSerializer.Serialize(new { Success = true, Message = message, FilePath = filePath, Format = format }, JsonOptions);
    }

    private static string FormatFileSize(long bytes)
    {
        string[] sizes = ["B", "KB", "MB", "GB"];
        int order = 0;
        double size = bytes;
        while (size >= 1024 && order < sizes.Length - 1)
        {
            order++;
            size /= 1024;
        }
        return $"{size:0.##} {sizes[order]}";
    }

    private static void EnsureDirectoryExists(string filePath)
    {
        var directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            Directory.CreateDirectory(directory);
    }

    #endregion
}
