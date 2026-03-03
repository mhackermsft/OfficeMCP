using OfficeMCP.Models;

namespace OfficeMCP.Services;

/// <summary>
/// Service interface for PowerPoint document operations.
/// </summary>
public interface IPowerPointDocumentService
{
    DocumentResult CreatePresentation(string filePath, string? title = null);
    DocumentResult AddSlide(string filePath, SlideLayoutOptions? layoutOptions = null);
    DocumentResult AddTitle(string filePath, int slideIndex, string title, string? subtitle = null, TextFormatting? textFormat = null);
    DocumentResult AddTextBox(string filePath, int slideIndex, string text, TextBoxOptions options);
    DocumentResult AddBulletPoints(string filePath, int slideIndex, string[] points, TextBoxOptions options);
    DocumentResult AddImage(string filePath, int slideIndex, string imagePath, long x, long y, ImageOptions? options = null);
    DocumentResult AddShape(string filePath, int slideIndex, ShapeOptions options);
    DocumentResult AddTable(string filePath, int slideIndex, string[][] data, long x, long y, long width, long height);
    DocumentResult SetSlideBackground(string filePath, int slideIndex, string color);
    DocumentResult DeleteSlide(string filePath, int slideIndex);
    DocumentResult DuplicateSlide(string filePath, int sourceIndex);
    DocumentResult ReorderSlide(string filePath, int fromIndex, int toIndex);
    ContentResult GetSlideText(string filePath, int slideIndex);
    ContentResult GetAllSlidesText(string filePath);
    ContentResult GetSlideCount(string filePath);
    DocumentResult AddSpeakerNotes(string filePath, int slideIndex, string notes);

    // Image extraction (returns base64 image data + context for AI OCR/captioning)
    IList<ImageExtractionResult> ExtractImages(string filePath);
}
