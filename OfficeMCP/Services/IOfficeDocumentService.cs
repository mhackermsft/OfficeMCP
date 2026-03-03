using OfficeMCP.Models;

namespace OfficeMCP.Services;

/// <summary>
/// Format-agnostic base interface for all document operations.
/// Implemented by format-specific services (Word, Excel, PowerPoint, PDF).
/// </summary>
public interface IOfficeDocumentService
{
    // Core document operations
    DocumentResult CreateDocument(string filePath, string? title = null, PageLayoutOptions? layout = null);
    ContentResult GetDocumentText(string filePath);
    DocumentResult AddMarkdownContent(string filePath, string markdown, string? baseImagePath = null);
    
    // Content elements
    DocumentResult AddParagraph(string filePath, string text, TextFormatting? textFormat = null, ParagraphFormatting? paragraphFormat = null);
    DocumentResult AddHeading(string filePath, string text, int level = 1, TextFormatting? textFormat = null);
    DocumentResult AddTable(string filePath, string[][] data, TableFormatting? tableFormat = null);
    DocumentResult AddImage(string filePath, string imagePath, ImageOptions? options = null);
    DocumentResult AddPageBreak(string filePath);
    DocumentResult AddBulletList(string filePath, string[] items, TextFormatting? textFormat = null);
    DocumentResult AddNumberedList(string filePath, string[] items, TextFormatting? textFormat = null);
    
    // Headers and footers
    DocumentResult AddHeader(string filePath, HeaderFooterOptions options);
    DocumentResult AddFooter(string filePath, HeaderFooterOptions options);
    
    // Layout and formatting
    DocumentResult SetPageLayout(string filePath, PageLayoutOptions options);
    
    // Reading and extraction
    ContentResult GetParagraphText(string filePath, int paragraphIndex);
    ContentResult GetParagraphRange(string filePath, int startIndex, int endIndex);
    
    // Conversion
    ContentResult ConvertToMarkdown(string filePath);

    // Image extraction (returns base64 image data + context for AI OCR/captioning)
    IList<ImageExtractionResult> ExtractImages(string filePath);
}
