using OfficeMCP.Models;

namespace OfficeMCP.Services;

/// <summary>
/// Service interface for Word document operations.
/// Inherits all format-agnostic operations from IOfficeDocumentService.
/// </summary>
public interface IWordDocumentService : IOfficeDocumentService
{
    /// <summary>
    /// Returns the document as an ordered list of content items (headings, paragraphs, images, tables)
    /// in reading order. Inline images appear immediately after their containing paragraph so that
    /// surrounding headings and paragraphs provide section context for AI OCR and captioning.
    /// Images include base64-encoded bytes and MimeType for direct AI vision analysis.
    /// </summary>
    IList<DocumentContentItem> GetRichContent(string filePath);
}
