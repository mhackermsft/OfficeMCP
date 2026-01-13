using OfficeMCP.Models;

namespace OfficeMCP.Services;

/// <summary>
/// Service interface for Word document operations.
/// </summary>
public interface IWordDocumentService
{
    DocumentResult CreateDocument(string filePath, string? title = null, PageLayoutOptions? layout = null);
    DocumentResult AddParagraph(string filePath, string text, TextFormatting? textFormat = null, ParagraphFormatting? paragraphFormat = null);
    DocumentResult AddHeading(string filePath, string text, int level = 1, TextFormatting? textFormat = null);
    DocumentResult AddTable(string filePath, string[][] data, TableFormatting? tableFormat = null);
    DocumentResult AddImage(string filePath, string imagePath, ImageOptions? options = null);
    DocumentResult AddHeader(string filePath, HeaderFooterOptions options);
    DocumentResult AddFooter(string filePath, HeaderFooterOptions options);
    DocumentResult AddPageBreak(string filePath);
    DocumentResult AddBulletList(string filePath, string[] items, TextFormatting? textFormat = null);
    DocumentResult AddNumberedList(string filePath, string[] items, TextFormatting? textFormat = null);
    ContentResult GetDocumentText(string filePath);
    ContentResult GetParagraphText(string filePath, int paragraphIndex);
    ContentResult GetParagraphRange(string filePath, int startIndex, int endIndex);
    DocumentResult SetPageLayout(string filePath, PageLayoutOptions options);
    DocumentResult AddMarkdownContent(string filePath, string markdown, string? baseImagePath = null);
    ContentResult ConvertToMarkdown(string filePath);
}
