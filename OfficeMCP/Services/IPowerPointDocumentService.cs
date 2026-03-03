using OfficeMCP.Models;

namespace OfficeMCP.Services;

/// <summary>
/// Service interface for PowerPoint document operations.
/// </summary>
public interface IPowerPointDocumentService
{
    // Core presentation operations
    DocumentResult CreatePresentation(string filePath, string? title = null);
    DocumentResult AddSlide(string filePath, SlideLayoutOptions? layoutOptions = null);
    DocumentResult DeleteSlide(string filePath, int slideIndex);
    DocumentResult DuplicateSlide(string filePath, int sourceIndex);
    DocumentResult ReorderSlide(string filePath, int fromIndex, int toIndex);
    DocumentResult SetSlideBackground(string filePath, int slideIndex, string color);
    DocumentResult SetSlideBackgroundGradient(string filePath, int slideIndex, GradientFillOptions gradient);
    DocumentResult SetSlideSize(string filePath, string size);

    // Text operations
    DocumentResult AddTitle(string filePath, int slideIndex, string title, string? subtitle = null, TextFormatting? textFormat = null);
    DocumentResult AddTextBox(string filePath, int slideIndex, string text, TextBoxOptions options);
    DocumentResult AddRichTextBox(string filePath, int slideIndex, TextBoxOptions options);
    DocumentResult AddBulletPoints(string filePath, int slideIndex, string[] points, TextBoxOptions options);

    // Shape operations
    DocumentResult AddShape(string filePath, int slideIndex, ShapeOptions options);
    DocumentResult AddLine(string filePath, int slideIndex, LineOptions options);
    DocumentResult AddConnector(string filePath, int slideIndex, ConnectorOptions options);
    DocumentResult AddGroupShape(string filePath, int slideIndex, long x, long y, long width, long height, GroupShapeItem[] items);

    // Media operations
    DocumentResult AddImage(string filePath, int slideIndex, string imagePath, long x, long y, ImageOptions? options = null);
    DocumentResult AddImageFromBase64(string filePath, int slideIndex, string base64Data, string mimeType, long x, long y, ImageOptions? options = null);
    DocumentResult AddTable(string filePath, int slideIndex, string[][] data, long x, long y, long width, long height);
    DocumentResult AddSpeakerNotes(string filePath, int slideIndex, string notes);

    // Z-order operations
    DocumentResult SetShapeZOrder(string filePath, int slideIndex, int shapeIndex, string position);
    DocumentResult ReorderShape(string filePath, int slideIndex, int fromIndex, int toIndex);

    // Read operations
    ContentResult GetSlideText(string filePath, int slideIndex);
    ContentResult GetAllSlidesText(string filePath);
    ContentResult GetSlideCount(string filePath);

    // Image extraction
    IList<ImageExtractionResult> ExtractImages(string filePath);
}
