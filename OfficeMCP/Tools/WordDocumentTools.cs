using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// MCP Tools for creating and manipulating Word documents.
/// </summary>
[McpServerToolType]
public sealed class WordDocumentTools(IWordDocumentService wordService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    [McpServerTool, Description("Create a new Word document (.docx). Optionally specify a title and page layout.")]
    public string CreateWordDocument(
        [Description("Full file path for the new document (e.g., C:/Documents/report.docx)")] string filePath,
        [Description("Optional document title to add as heading")] string? title = null,
        [Description("Page orientation: 'Portrait' or 'Landscape'")] string orientation = "Portrait",
        [Description("Page size: 'Letter', 'Legal', 'A4', or 'A3'")] string pageSize = "Letter",
        [Description("Top margin in inches")] double marginTop = 1.0,
        [Description("Bottom margin in inches")] double marginBottom = 1.0,
        [Description("Left margin in inches")] double marginLeft = 1.0,
        [Description("Right margin in inches")] double marginRight = 1.0)
    {
        var layout = new PageLayoutOptions(
            Orientation: orientation,
            PageSize: pageSize,
            MarginTop: marginTop,
            MarginBottom: marginBottom,
            MarginLeft: marginLeft,
            MarginRight: marginRight
        );
        
        var result = wordService.CreateDocument(filePath, title, layout);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a paragraph of text to an existing Word document with optional formatting.")]
    public string AddParagraphToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("The text content to add")] string text,
        [Description("Make text bold")] bool bold = false,
        [Description("Make text italic")] bool italic = false,
        [Description("Underline the text")] bool underline = false,
        [Description("Font name (e.g., 'Arial', 'Times New Roman')")] string? fontName = null,
        [Description("Font size in points (e.g., 12)")] int? fontSize = null,
        [Description("Font color as hex (e.g., '0000FF' for blue)")] string? fontColor = null,
        [Description("Paragraph alignment: 'Left', 'Center', 'Right', or 'Justify'")] string alignment = "Left",
        [Description("Line spacing multiplier (e.g., 1.5 for 1.5x spacing)")] double? lineSpacing = null)
    {
        var textFormat = new TextFormatting(bold, italic, underline, false, fontName, fontSize, fontColor);
        var paragraphFormat = new ParagraphFormatting(alignment, lineSpacing);
        
        var result = wordService.AddParagraph(filePath, text, textFormat, paragraphFormat);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a heading to a Word document. Headings use larger, bold text and support levels 1-9.")]
    public string AddHeadingToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("The heading text")] string text,
        [Description("Heading level from 1 (largest) to 9 (smallest)")] int level = 1)
    {
        var result = wordService.AddHeading(filePath, text, level);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a table to a Word document. Provide data as a 2D array where each inner array is a row.")]
    public string AddTableToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("Table data as JSON 2D array, e.g., [[\"Header1\",\"Header2\"],[\"Row1Col1\",\"Row1Col2\"]]")] string tableDataJson,
        [Description("First row is header with special formatting")] bool hasHeader = true,
        [Description("Border color as hex (e.g., '000000' for black)")] string? borderColor = null,
        [Description("Border width in points")] double borderWidth = 1.0,
        [Description("Header row background color as hex")] string? headerBackgroundColor = null,
        [Description("Alternate row background color as hex for striped effect")] string? alternateRowColor = null)
    {
        try
        {
            var tableData = JsonSerializer.Deserialize<string[][]>(tableDataJson, JsonOptions);
            if (tableData == null || tableData.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty table data"), JsonOptions);
            }

            var tableFormat = new TableFormatting(borderColor, borderWidth, hasHeader, headerBackgroundColor, alternateRowColor);
            var result = wordService.AddTable(filePath, tableData, tableFormat);
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description("Add an image to a Word document.")]
    public string AddImageToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("Full path to the image file (supports jpg, png, gif, bmp, tiff)")] string imagePath,
        [Description("Image width in inches")] double widthInches = 4.0,
        [Description("Image height in inches")] double heightInches = 3.0,
        [Description("Alt text description for accessibility")] string? altText = null)
    {
        var options = new ImageOptions(
            WidthEmu: (long)(widthInches * 914400),
            HeightEmu: (long)(heightInches * 914400),
            AltText: altText
        );
        
        var result = wordService.AddImage(filePath, imagePath, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a header to all pages in a Word document.")]
    public string AddHeaderToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("Text for left side of header")] string? leftContent = null,
        [Description("Text for center of header")] string? centerContent = null,
        [Description("Text for right side of header")] string? rightContent = null,
        [Description("Include automatic page numbering")] bool includePageNumber = false,
        [Description("Include current date")] bool includeDate = false)
    {
        var options = new HeaderFooterOptions(leftContent, centerContent, rightContent, includePageNumber, includeDate);
        var result = wordService.AddHeader(filePath, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a footer to all pages in a Word document.")]
    public string AddFooterToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("Text for left side of footer")] string? leftContent = null,
        [Description("Text for center of footer")] string? centerContent = null,
        [Description("Text for right side of footer")] string? rightContent = null,
        [Description("Include automatic page numbering")] bool includePageNumber = true,
        [Description("Include current date")] bool includeDate = false)
    {
        var options = new HeaderFooterOptions(leftContent, centerContent, rightContent, includePageNumber, includeDate);
        var result = wordService.AddFooter(filePath, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a page break to move subsequent content to a new page.")]
    public string AddPageBreakToWord(
        [Description("Path to the existing Word document")] string filePath)
    {
        var result = wordService.AddPageBreak(filePath);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a bullet list to a Word document.")]
    public string AddBulletListToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("List items as JSON array, e.g., [\"Item 1\", \"Item 2\", \"Item 3\"]")] string itemsJson,
        [Description("Make text bold")] bool bold = false,
        [Description("Font size in points")] int? fontSize = null)
    {
        try
        {
            var items = JsonSerializer.Deserialize<string[]>(itemsJson, JsonOptions);
            if (items == null || items.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty items array"), JsonOptions);
            }

            var textFormat = new TextFormatting(Bold: bold, FontSize: fontSize);
            var result = wordService.AddBulletList(filePath, items, textFormat);
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description("Add a numbered list to a Word document.")]
    public string AddNumberedListToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("List items as JSON array, e.g., [\"Step 1\", \"Step 2\", \"Step 3\"]")] string itemsJson,
        [Description("Make text bold")] bool bold = false,
        [Description("Font size in points")] int? fontSize = null)
    {
        try
        {
            var items = JsonSerializer.Deserialize<string[]>(itemsJson, JsonOptions);
            if (items == null || items.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty items array"), JsonOptions);
            }

            var textFormat = new TextFormatting(Bold: bold, FontSize: fontSize);
            var result = wordService.AddNumberedList(filePath, items, textFormat);
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description("Get all text content from a Word document.")]
    public string GetWordDocumentText(
        [Description("Path to the Word document")] string filePath)
    {
        var result = wordService.GetDocumentText(filePath);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Get text from a specific paragraph by index (0-based).")]
    public string GetWordParagraphText(
        [Description("Path to the Word document")] string filePath,
        [Description("Paragraph index (0-based)")] int paragraphIndex)
    {
        var result = wordService.GetParagraphText(filePath, paragraphIndex);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Get text from a range of paragraphs.")]
    public string GetWordParagraphRange(
        [Description("Path to the Word document")] string filePath,
        [Description("Starting paragraph index (0-based, inclusive)")] int startIndex,
        [Description("Ending paragraph index (0-based, inclusive)")] int endIndex)
    {
        var result = wordService.GetParagraphRange(filePath, startIndex, endIndex);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Change the page layout settings of an existing Word document.")]
    public string SetWordPageLayout(
        [Description("Path to the existing Word document")] string filePath,
        [Description("Page orientation: 'Portrait' or 'Landscape'")] string orientation = "Portrait",
        [Description("Page size: 'Letter', 'Legal', 'A4', or 'A3'")] string pageSize = "Letter",
        [Description("Top margin in inches")] double marginTop = 1.0,
        [Description("Bottom margin in inches")] double marginBottom = 1.0,
        [Description("Left margin in inches")] double marginLeft = 1.0,
        [Description("Right margin in inches")] double marginRight = 1.0)
    {
        var layout = new PageLayoutOptions(orientation, marginTop, marginBottom, marginLeft, marginRight, pageSize);
        var result = wordService.SetPageLayout(filePath, layout);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Add Markdown content to a Word document with full formatting preserved. 
Supports: headings (# to ######), **bold**, *italic*, ***bold italic***, ~~strikethrough~~, `inline code`, 
```code blocks```, bullet lists (- or *), numbered lists (1.), > blockquotes, tables (|col1|col2|), 
horizontal rules (---), [links](url), and ![images](path). Images can use relative paths with baseImagePath.")]
    public string AddMarkdownToWord(
        [Description("Path to the existing Word document")] string filePath,
        [Description("Markdown formatted text content")] string markdown,
        [Description("Base path for resolving relative image paths (optional)")] string? baseImagePath = null)
    {
        var result = wordService.AddMarkdownContent(filePath, markdown, baseImagePath);
        return JsonSerializer.Serialize(result, JsonOptions);
    }
}
