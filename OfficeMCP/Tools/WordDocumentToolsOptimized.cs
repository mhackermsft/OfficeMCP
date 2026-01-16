using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// LEGACY: Format-specific Word tools - kept for backward compatibility.
/// Use the unified office_* tools from OfficeDocumentToolsConsolidated instead.
/// This class is no longer registered as an MCP tool provider.
/// </summary>
// [McpServerToolType] - Disabled: Use consolidated office_* tools instead
public sealed class WordDocumentToolsOptimized(IWordDocumentService wordService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    #region Core Document Operations

    // [McpServerTool] - Disabled: Use office_create instead
    public string CreateWordDocument(
        [Description("Full path (e.g., C:/docs/report.docx)")] string filePath,
        [Description("Document title as Heading 1")] string? title = null,
        [Description("Markdown content for rich formatting")] string? markdown = null,
        [Description("Base path for relative image paths")] string? baseImagePath = null,
        [Description("Portrait or Landscape")] string orientation = "Portrait",
        [Description("Letter, Legal, A4, or A3")] string pageSize = "Letter")
    {
        var layout = new PageLayoutOptions(Orientation: orientation, PageSize: pageSize);
        var result = wordService.CreateDocument(filePath, title, layout);

        if (!result.Success)
            return JsonSerializer.Serialize(result, JsonOptions);

        if (!string.IsNullOrWhiteSpace(markdown))
        {
            var mdResult = wordService.AddMarkdownContent(filePath, markdown, baseImagePath);
            if (!mdResult.Success)
                return JsonSerializer.Serialize(new DocumentResult(false, $"Created but markdown failed: {mdResult.Message}", filePath), JsonOptions);
            return JsonSerializer.Serialize(new DocumentResult(true, "Document created with content", filePath), JsonOptions);
        }

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "word_read", Destructive = false, ReadOnly = true), Description("Reads text from a Word document. Returns all text by default, or specific paragraphs by index/range.")]
    public string ReadWordDocument(
        [Description("Path to the Word document")] string filePath,
        [Description("all (default), paragraph, or range")] string readType = "all",
        [Description("Paragraph index (0-based) for 'paragraph' mode")] int? paragraphIndex = null,
        [Description("Start index (0-based) for 'range' mode")] int? startIndex = null,
        [Description("End index (0-based) for 'range' mode")] int? endIndex = null)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new ContentResult(false, null, $"File not found: {filePath}. Use word_create to create a document first."), JsonOptions);

        ContentResult result = readType.ToLowerInvariant() switch
        {
            "paragraph" when paragraphIndex.HasValue => wordService.GetParagraphText(filePath, paragraphIndex.Value),
            "paragraph" => new ContentResult(false, null, "paragraphIndex is required when readType is 'paragraph'"),
            "range" when startIndex.HasValue && endIndex.HasValue => wordService.GetParagraphRange(filePath, startIndex.Value, endIndex.Value),
            "range" => new ContentResult(false, null, "startIndex and endIndex are required when readType is 'range'"),
            _ => wordService.GetDocumentText(filePath)
        };

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "word_add_content", Destructive = false, ReadOnly = false), Description("Adds markdown content to an existing Word document. Supports # headings, **bold**, *italic*, - lists, 1. numbered, | tables |, ```code```, [links](url), ![images](path).")]
    public string AddMarkdownToWord(
        [Description("Path to the Word document")] string filePath,
        [Description("Markdown content to add")] string markdown,
        [Description("Base path for relative image paths")] string? baseImagePath = null)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            return JsonSerializer.Serialize(new DocumentResult(false, "Markdown content is required", Suggestion: "Provide markdown text in the 'markdown' parameter"), JsonOptions);
        
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {filePath}", Suggestion: "Use word_create to create the document first, or verify the file path"), JsonOptions);
        
        var result = wordService.AddMarkdownContent(filePath, markdown, baseImagePath);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    #endregion

    #region Document Elements

    [McpServerTool(Name = "word_add_element", Destructive = false, ReadOnly = false), Description("Adds page break, header, or footer to a Word document.")]
    public string AddWordDocumentElement(
        [Description("Path to the Word document")] string filePath,
        [Description("pageBreak, header, or footer")] string operation,
        [Description("Left content for header/footer")] string? leftContent = null,
        [Description("Center content for header/footer")] string? centerContent = null,
        [Description("Right content for header/footer")] string? rightContent = null,
        [Description("Include page number")] bool includePageNumber = false,
        [Description("Include date")] bool includeDate = false)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {filePath}", Suggestion: "Use word_create to create the document first"), JsonOptions);

        DocumentResult result = operation.ToLowerInvariant() switch
        {
            "pagebreak" => wordService.AddPageBreak(filePath),
            "header" => wordService.AddHeader(filePath, new HeaderFooterOptions(leftContent, centerContent, rightContent, includePageNumber, includeDate)),
            "footer" => wordService.AddFooter(filePath, new HeaderFooterOptions(leftContent, centerContent, rightContent, includePageNumber, includeDate)),
            _ => new DocumentResult(false, $"Unknown operation: '{operation}'", Suggestion: "Valid operations: pageBreak, header, footer")
        };
        
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "word_add_image", Destructive = false, ReadOnly = false), Description("Adds an image to a Word document.")]
    public string AddImageToWord(
        [Description("Path to the Word document")] string filePath,
        [Description("Full path to image (jpg, png, gif, bmp, tiff)")] string imagePath,
        [Description("Width in inches")] double widthInches = 4.0,
        [Description("Height in inches")] double heightInches = 3.0,
        [Description("Alt text for accessibility")] string? altText = null)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"Document not found: {filePath}", Suggestion: "Use word_create to create the document first"), JsonOptions);
        
        if (!File.Exists(imagePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"Image not found: {imagePath}", Suggestion: "Verify the image path exists and is accessible"), JsonOptions);

        var options = new ImageOptions(
            WidthEmu: (long)(widthInches * 914400),
            HeightEmu: (long)(heightInches * 914400),
            AltText: altText
        );
        var result = wordService.AddImage(filePath, imagePath, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    #endregion

    #region Conversion Tools

    [McpServerTool(Name = "word_convert", Destructive = false, ReadOnly = false), Description("Converts between Word and Markdown formats. Use 'direction' to specify: word_to_md, md_to_word, or word_to_md_file.")]
    public string ConvertDocument(
        [Description("Path to source file")] string sourcePath,
        [Description("word_to_md (returns markdown), md_to_word (creates .docx), or word_to_md_file (creates .md file)")] string direction,
        [Description("Output path (optional, auto-generated if not provided)")] string? outputPath = null,
        [Description("Base path for relative images")] string? baseImagePath = null,
        [Description("Portrait or Landscape (for md_to_word)")] string orientation = "Portrait",
        [Description("Letter, Legal, A4, A3 (for md_to_word)")] string pageSize = "Letter")
    {
        try
        {
            return direction.ToLowerInvariant() switch
            {
                "word_to_md" => JsonSerializer.Serialize(wordService.ConvertToMarkdown(sourcePath), JsonOptions),
                "md_to_word" => ConvertMarkdownToWord(sourcePath, outputPath, baseImagePath, orientation, pageSize),
                "word_to_md_file" => ConvertWordToMarkdownFile(sourcePath, outputPath),
                _ => JsonSerializer.Serialize(new DocumentResult(false, $"Unknown direction: '{direction}'", Suggestion: "Valid directions: word_to_md, md_to_word, word_to_md_file"), JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Conversion error: {ex.Message}", Suggestion: "Verify source file exists and is not corrupted"), JsonOptions);
        }
    }

    private string ConvertMarkdownToWord(string markdownFilePath, string? outputPath, string? baseImagePath, string orientation, string pageSize)
    {
        if (!File.Exists(markdownFilePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {markdownFilePath}", Suggestion: "Verify the markdown file path is correct"), JsonOptions);

        var markdown = File.ReadAllText(markdownFilePath);
        if (string.IsNullOrWhiteSpace(markdown))
            return JsonSerializer.Serialize(new DocumentResult(false, "Markdown file is empty", Suggestion: "Add content to the markdown file before converting"), JsonOptions);

        var docxPath = outputPath ?? Path.ChangeExtension(markdownFilePath, ".docx");
        var effectiveBaseImagePath = baseImagePath ?? Path.GetDirectoryName(markdownFilePath);

        var layout = new PageLayoutOptions(Orientation: orientation, PageSize: pageSize);
        var createResult = wordService.CreateDocument(docxPath, null, layout);
        if (!createResult.Success)
            return JsonSerializer.Serialize(createResult, JsonOptions);

        var mdResult = wordService.AddMarkdownContent(docxPath, markdown, effectiveBaseImagePath);
        if (!mdResult.Success)
            return JsonSerializer.Serialize(new DocumentResult(false, $"Created but conversion failed: {mdResult.Message}", docxPath, Suggestion: "Check markdown syntax is valid"), JsonOptions);

        return JsonSerializer.Serialize(new DocumentResult(true, $"Converted to '{docxPath}'", docxPath), JsonOptions);
    }

    private string ConvertWordToMarkdownFile(string wordFilePath, string? outputPath)
    {
        if (!File.Exists(wordFilePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {wordFilePath}", Suggestion: "Verify the Word document path is correct"), JsonOptions);

        var conversionResult = wordService.ConvertToMarkdown(wordFilePath);
        if (!conversionResult.Success || string.IsNullOrEmpty(conversionResult.Content))
            return JsonSerializer.Serialize(new DocumentResult(false, conversionResult.ErrorMessage ?? "Conversion failed", Suggestion: "Ensure the file is a valid .docx document"), JsonOptions);

        var mdPath = outputPath ?? Path.ChangeExtension(wordFilePath, ".md");
        var directory = Path.GetDirectoryName(mdPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            Directory.CreateDirectory(directory);

        File.WriteAllText(mdPath, conversionResult.Content);
        return JsonSerializer.Serialize(new DocumentResult(true, $"Converted to '{mdPath}'", mdPath), JsonOptions);
    }

    #endregion

    #region Batch Operations

    [McpServerTool(Name = "word_batch", Destructive = false, ReadOnly = false), Description("Performs multiple operations on a Word document. Operations: markdown, heading, paragraph, bulletList, numberedList, table, image, pageBreak, header, footer.")]
    public string BatchModifyWordDocument(
        [Description("Path to the Word document")] string filePath,
        [Description("JSON array: [{\"type\":\"heading\",\"text\":\"Title\",\"level\":1}, {\"type\":\"paragraph\",\"text\":\"Content\"}]")] string operationsJson)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {filePath}", Suggestion: "Use word_create to create the document first"), JsonOptions);

        try
        {
            var operations = JsonSerializer.Deserialize<WordOperation[]>(operationsJson, JsonOptions);
            if (operations == null || operations.Length == 0)
                return JsonSerializer.Serialize(new DocumentResult(false, "No operations provided", Suggestion: "Provide a JSON array of operations, e.g., [{\"type\":\"heading\",\"text\":\"Title\"}]"), JsonOptions);

            var details = new List<OperationOutcome>();
            int successCount = 0, failCount = 0;

            for (int i = 0; i < operations.Length; i++)
            {
                var op = operations[i];
                try
                {
                    var opResult = op.Type.ToLowerInvariant() switch
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
                        _ => new DocumentResult(false, $"Unknown type: '{op.Type}'", Suggestion: "Valid types: markdown, heading, paragraph, bulletList, numberedList, table, image, pageBreak, header, footer")
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

            return JsonSerializer.Serialize(new BatchOperationResult(
                Success: failCount == 0,
                Message: failCount == 0 ? $"All {successCount} operations completed" : $"{successCount} succeeded, {failCount} failed",
                TotalOperations: operations.Length,
                SuccessfulOperations: successCount,
                FailedOperations: failCount,
                Details: details
            ), JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON: {ex.Message}", Suggestion: "Ensure operationsJson is a valid JSON array"), JsonOptions);
        }
    }

    #endregion

    #region Private Operation Processors

    private DocumentResult ProcessMarkdown(string filePath, WordOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.Markdown))
            return new DocumentResult(false, "Markdown content is required");
        return wordService.AddMarkdownContent(filePath, op.Markdown, op.BaseImagePath);
    }

    private DocumentResult ProcessHeading(string filePath, WordOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.Text))
            return new DocumentResult(false, "Heading text is required");
        return wordService.AddHeading(filePath, op.Text, op.Level ?? 1);
    }

    private DocumentResult ProcessParagraph(string filePath, WordOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.Text))
            return new DocumentResult(false, "Paragraph text is required");
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
            return new DocumentResult(false, "Items array is required");
        var textFormat = new TextFormatting(Bold: op.Bold ?? false, FontSize: op.FontSize);
        return wordService.AddBulletList(filePath, op.Items, textFormat);
    }

    private DocumentResult ProcessNumberedList(string filePath, WordOperation op)
    {
        if (op.Items == null || op.Items.Length == 0)
            return new DocumentResult(false, "Items array is required");
        var textFormat = new TextFormatting(Bold: op.Bold ?? false, FontSize: op.FontSize);
        return wordService.AddNumberedList(filePath, op.Items, textFormat);
    }

    private DocumentResult ProcessTable(string filePath, WordOperation op)
    {
        if (op.TableData == null || op.TableData.Length == 0)
            return new DocumentResult(false, "Table data is required");
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
            return new DocumentResult(false, "Image path is required");
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
