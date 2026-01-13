using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// AI-Optimized MCP Tools for Word documents. Reduces tool calls through batch operations and markdown support.
/// </summary>
[McpServerToolType]
public sealed class WordDocumentToolsOptimized(IWordDocumentService wordService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    [McpServerTool, Description(@"Create a new Word document and optionally add content in a single call.
Use this to create documents with initial content. For best results, use the 'markdown' parameter to add rich content.

**Recommended**: Use markdown parameter for complex content - it's the most efficient way to add formatted text, headings, lists, tables, and more.

**Examples**:
- Simple: {""filePath"": ""C:/docs/report.docx"", ""title"": ""Quarterly Report""}
- With markdown: {""filePath"": ""C:/docs/report.docx"", ""markdown"": ""# Title\n\nIntro paragraph.\n\n## Section 1\n\n- Bullet 1\n- Bullet 2""}

**Supported markdown**: # headings, **bold**, *italic*, - bullet lists, 1. numbered lists, > blockquotes, | tables |, ``` code blocks, [links](url), ![images](path)")]
    public string CreateWordDocument(
        [Description("Full file path for the new document (e.g., C:/Documents/report.docx)")] string filePath,
        [Description("Optional document title as heading")] string? title = null,
        [Description("Markdown content to add (recommended for rich formatting)")] string? markdown = null,
        [Description("Base path for resolving relative image paths in markdown")] string? baseImagePath = null,
        [Description("Page orientation: 'Portrait' or 'Landscape'")] string orientation = "Portrait",
        [Description("Page size: 'Letter', 'Legal', 'A4', or 'A3'")] string pageSize = "Letter")
    {
        var layout = new PageLayoutOptions(Orientation: orientation, PageSize: pageSize);
        var result = wordService.CreateDocument(filePath, title, layout);

        if (!result.Success)
        {
            return JsonSerializer.Serialize(result, JsonOptions);
        }

        // If markdown is provided, add it to the document
        if (!string.IsNullOrWhiteSpace(markdown))
        {
            var mdResult = wordService.AddMarkdownContent(filePath, markdown, baseImagePath);
            if (!mdResult.Success)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, $"Document created but markdown failed: {mdResult.Message}", filePath), JsonOptions);
            }
            return JsonSerializer.Serialize(new DocumentResult(true, "Document created with markdown content", filePath), JsonOptions);
        }

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Convert a markdown file (.md) to a Word document (.docx). The output file will have the same name and path but with a .docx extension.

**Example**: Converting 'C:/docs/report.md' creates 'C:/docs/report.docx'

All markdown formatting is preserved including headings, bold, italic, lists, tables, code blocks, links, and images.")]
    public string ConvertMarkdownFileToWord(
        [Description("Full path to the markdown file (e.g., C:/Documents/report.md)")] string markdownFilePath,
        [Description("Optional: Custom output path for the Word document. If not provided, uses same path with .docx extension")] string? outputPath = null,
        [Description("Base path for resolving relative image paths in the markdown")] string? baseImagePath = null,
        [Description("Page orientation: 'Portrait' or 'Landscape'")] string orientation = "Portrait",
        [Description("Page size: 'Letter', 'Legal', 'A4', or 'A3'")] string pageSize = "Letter")
    {
        try
        {
            // Validate input file exists
            if (!File.Exists(markdownFilePath))
            {
                return JsonSerializer.Serialize(new DocumentResult(false, $"Markdown file not found: {markdownFilePath}"), JsonOptions);
            }

            // Read markdown content
            var markdown = File.ReadAllText(markdownFilePath);
            if (string.IsNullOrWhiteSpace(markdown))
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "Markdown file is empty"), JsonOptions);
            }

            // Determine output path
            var docxPath = outputPath ?? Path.ChangeExtension(markdownFilePath, ".docx");

            // Use the markdown file's directory as base image path if not specified
            var effectiveBaseImagePath = baseImagePath ?? Path.GetDirectoryName(markdownFilePath);

            // Create the Word document
            var layout = new PageLayoutOptions(Orientation: orientation, PageSize: pageSize);
            var createResult = wordService.CreateDocument(docxPath, null, layout);
            if (!createResult.Success)
            {
                return JsonSerializer.Serialize(createResult, JsonOptions);
            }

            // Add the markdown content
            var mdResult = wordService.AddMarkdownContent(docxPath, markdown, effectiveBaseImagePath);
            if (!mdResult.Success)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, $"Document created but markdown conversion failed: {mdResult.Message}", docxPath), JsonOptions);
            }

            return JsonSerializer.Serialize(new DocumentResult(true, $"Successfully converted '{markdownFilePath}' to '{docxPath}'", docxPath), JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Error converting markdown file: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description(@"Add markdown content to an existing Word document. This is the simplest and most efficient way to add rich content.

**Supported markdown**: # headings (1-6), **bold**, *italic*, ***bold italic***, ~~strikethrough~~, `inline code`, 
```code blocks```, - bullet lists, 1. numbered lists, > blockquotes, | tables |, [links](url), ![images](path), --- horizontal rules

**Example usage**: Just pass the markdown text directly - no JSON escaping needed!")]
    public string AddMarkdownToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("Markdown content to add - supports headings, bold, italic, lists, tables, code blocks, links, images, and more")] string markdown,
        [Description("Base path for resolving relative image paths in markdown")] string? baseImagePath = null)
    {
        if (string.IsNullOrWhiteSpace(markdown))
        {
            return JsonSerializer.Serialize(new DocumentResult(false, "Markdown content is required"), JsonOptions);
        }
        
        var result = wordService.AddMarkdownContent(filePath, markdown, baseImagePath);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Add a page break, header, or footer to a Word document.

**Use 'operation' parameter**:
- 'pageBreak': Insert a page break
- 'header': Set page header (use leftContent, centerContent, rightContent)
- 'footer': Set page footer (use leftContent, centerContent, rightContent)")]
    public string AddWordDocumentElement(
        [Description("Path to the existing Word document")] string filePath,
        [Description("Operation: 'pageBreak', 'header', or 'footer'")] string operation,
        [Description("Content for left side of header/footer")] string? leftContent = null,
        [Description("Content for center of header/footer")] string? centerContent = null,
        [Description("Content for right side of header/footer")] string? rightContent = null,
        [Description("Include page number in header/footer")] bool includePageNumber = false,
        [Description("Include date in header/footer")] bool includeDate = false)
    {
        DocumentResult result = operation.ToLowerInvariant() switch
        {
            "pagebreak" => wordService.AddPageBreak(filePath),
            "header" => wordService.AddHeader(filePath, new HeaderFooterOptions(leftContent, centerContent, rightContent, includePageNumber, includeDate)),
            "footer" => wordService.AddFooter(filePath, new HeaderFooterOptions(leftContent, centerContent, rightContent, includePageNumber, includeDate)),
            _ => new DocumentResult(false, $"Unknown operation: {operation}. Use 'pageBreak', 'header', or 'footer'.")
        };
        
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Add an image to a Word document.")]
    public string AddImageToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("Full path to the image file (supports jpg, png, gif, bmp, tiff)")] string imagePath,
        [Description("Image width in inches (default: 4)")] double widthInches = 4.0,
        [Description("Image height in inches (default: 3)")] double heightInches = 3.0,
        [Description("Alt text for accessibility")] string? altText = null)
    {
        var options = new ImageOptions(
            WidthEmu: (long)(widthInches * 914400),
            HeightEmu: (long)(heightInches * 914400),
            AltText: altText
        );
        var result = wordService.AddImage(filePath, imagePath, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Perform batch operations on a Word document using a JSON array. Use this for complex multi-step modifications.

**Operations JSON array** - each object has a 'type' and type-specific properties:

| Type | Key Properties |
|------|----------------|
| markdown | markdown, baseImagePath |
| heading | text, level (1-9) |
| paragraph | text, bold, italic, alignment |
| bulletList | items (array of strings) |
| numberedList | items (array of strings) |
| table | tableData (2D array), hasHeader |
| image | imagePath, widthInches, heightInches |
| pageBreak | (none) |
| header | leftContent, centerContent, rightContent, includePageNumber |
| footer | leftContent, centerContent, rightContent, includePageNumber |")]
    public string BatchModifyWordDocument(
        [Description("Path to the existing Word document")] string filePath,
        [Description("JSON array of operations")] string operationsJson)
    {
        try
        {
            var operations = JsonSerializer.Deserialize<WordOperation[]>(operationsJson, JsonOptions);
            if (operations == null || operations.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "No operations provided"), JsonOptions);
            }

            var details = new List<OperationOutcome>();
            int successCount = 0;
            int failCount = 0;

            for (int i = 0; i < operations.Length; i++)
            {
                var op = operations[i];
                DocumentResult? opResult = null;

                try
                {
                    opResult = op.Type.ToLowerInvariant() switch
                    {
                        "markdown" => ProcessMarkdown(filePath, op),
                        "heading" => ProcessHeading(filePath, op),
                        "paragraph" => ProcessParagraph(filePath, op),
                        "bulletlist" => ProcessBulletList(filePath, op),
                        "numberedlist" => ProcessNumberedList(filePath, op),
                        "table" => ProcessTable(filePath, op),
                        "image" => ProcessImage(filePath, op),
                        "pagebreak" => wordService.AddPageBreak(filePath),
                        "header" => ProcessHeader(filePath, op),
                        "footer" => ProcessFooter(filePath, op),
                        _ => new DocumentResult(false, $"Unknown operation type: {op.Type}")
                    };

                    if (opResult.Success)
                    {
                        successCount++;
                        details.Add(new OperationOutcome(i, op.Type, true, "Success"));
                    }
                    else
                    {
                        failCount++;
                        details.Add(new OperationOutcome(i, op.Type, false, opResult.Message));
                    }
                }
                catch (Exception ex)
                {
                    failCount++;
                    details.Add(new OperationOutcome(i, op.Type, false, ex.Message));
                }
            }

            var batchResult = new BatchOperationResult(
                Success: failCount == 0,
                Message: failCount == 0
                    ? $"All {successCount} operations completed successfully"
                    : $"{successCount} succeeded, {failCount} failed",
                TotalOperations: operations.Length,
                SuccessfulOperations: successCount,
                FailedOperations: failCount,
                Details: details
            );

            return JsonSerializer.Serialize(batchResult, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description(@"Read content from a Word document.

**Options for 'readType' parameter**:
- 'all' (default): Get all text content from the document
- 'paragraph': Get specific paragraph by index (use paragraphIndex)
- 'range': Get range of paragraphs (use startIndex and endIndex)

**Examples**:
- Get all text: {""filePath"": ""C:/docs/report.docx""}
- Get paragraph 3: {""filePath"": ""C:/docs/report.docx"", ""readType"": ""paragraph"", ""paragraphIndex"": 3}
- Get paragraphs 5-10: {""filePath"": ""C:/docs/report.docx"", ""readType"": ""range"", ""startIndex"": 5, ""endIndex"": 10}")]
    public string ReadWordDocument(
        [Description("Path to the Word document")] string filePath,
        [Description("Type of read: 'all', 'paragraph', or 'range'")] string readType = "all",
        [Description("Paragraph index for 'paragraph' read type (0-based)")] int? paragraphIndex = null,
        [Description("Start index for 'range' read type (0-based)")] int? startIndex = null,
        [Description("End index for 'range' read type (0-based)")] int? endIndex = null)
    {
        ContentResult result = readType.ToLowerInvariant() switch
        {
            "paragraph" when paragraphIndex.HasValue => wordService.GetParagraphText(filePath, paragraphIndex.Value),
            "range" when startIndex.HasValue && endIndex.HasValue => wordService.GetParagraphRange(filePath, startIndex.Value, endIndex.Value),
            _ => wordService.GetDocumentText(filePath)
        };

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    #region Private Operation Processors

    private DocumentResult ProcessMarkdown(string filePath, WordOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.Markdown))
        {
            return new DocumentResult(false, "Markdown content is required");
        }
        return wordService.AddMarkdownContent(filePath, op.Markdown, op.BaseImagePath);
    }

    private DocumentResult ProcessHeading(string filePath, WordOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.Text))
        {
            return new DocumentResult(false, "Heading text is required");
        }
        return wordService.AddHeading(filePath, op.Text, op.Level ?? 1);
    }

    private DocumentResult ProcessParagraph(string filePath, WordOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.Text))
        {
            return new DocumentResult(false, "Paragraph text is required");
        }
        var textFormat = new TextFormatting(
            Bold: op.Bold ?? false,
            Italic: op.Italic ?? false,
            Underline: op.Underline ?? false,
            FontName: op.FontName,
            FontSize: op.FontSize,
            FontColor: op.FontColor
        );
        var paragraphFormat = new ParagraphFormatting(
            Alignment: op.Alignment ?? "Left",
            LineSpacing: op.LineSpacing
        );
        return wordService.AddParagraph(filePath, op.Text, textFormat, paragraphFormat);
    }

    private DocumentResult ProcessBulletList(string filePath, WordOperation op)
    {
        if (op.Items == null || op.Items.Length == 0)
        {
            return new DocumentResult(false, "Items array is required for bullet list");
        }
        var textFormat = new TextFormatting(Bold: op.Bold ?? false, FontSize: op.FontSize);
        return wordService.AddBulletList(filePath, op.Items, textFormat);
    }

    private DocumentResult ProcessNumberedList(string filePath, WordOperation op)
    {
        if (op.Items == null || op.Items.Length == 0)
        {
            return new DocumentResult(false, "Items array is required for numbered list");
        }
        var textFormat = new TextFormatting(Bold: op.Bold ?? false, FontSize: op.FontSize);
        return wordService.AddNumberedList(filePath, op.Items, textFormat);
    }

    private DocumentResult ProcessTable(string filePath, WordOperation op)
    {
        if (op.TableData == null || op.TableData.Length == 0)
        {
            return new DocumentResult(false, "Table data is required");
        }
        var tableFormat = new TableFormatting(
            BorderColor: op.BorderColor,
            BorderWidth: op.BorderWidth ?? 1.0,
            HasHeader: op.HasHeader ?? true,
            HeaderBackgroundColor: op.HeaderBackgroundColor,
            AlternateRowColor: op.AlternateRowColor
        );
        return wordService.AddTable(filePath, op.TableData, tableFormat);
    }

    private DocumentResult ProcessImage(string filePath, WordOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.ImagePath))
        {
            return new DocumentResult(false, "Image path is required");
        }
        var options = new ImageOptions(
            WidthEmu: (long)((op.WidthInches ?? 4.0) * 914400),
            HeightEmu: (long)((op.HeightInches ?? 3.0) * 914400),
            AltText: op.AltText
        );
        return wordService.AddImage(filePath, op.ImagePath, options);
    }

    private DocumentResult ProcessHeader(string filePath, WordOperation op)
    {
        var options = new HeaderFooterOptions(
            LeftContent: op.LeftContent,
            CenterContent: op.CenterContent,
            RightContent: op.RightContent,
            IncludePageNumber: op.IncludePageNumber ?? false,
            IncludeDate: op.IncludeDate ?? false
        );
        return wordService.AddHeader(filePath, options);
    }

    private DocumentResult ProcessFooter(string filePath, WordOperation op)
    {
        var options = new HeaderFooterOptions(
            LeftContent: op.LeftContent,
            CenterContent: op.CenterContent,
            RightContent: op.RightContent,
            IncludePageNumber: op.IncludePageNumber ?? true,
            IncludeDate: op.IncludeDate ?? false
        );
        return wordService.AddFooter(filePath, options);
    }

    #endregion
}
