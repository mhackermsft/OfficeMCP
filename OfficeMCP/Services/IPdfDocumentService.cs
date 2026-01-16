using OfficeMCP.Models;

namespace OfficeMCP.Services;

/// <summary>
/// Service interface for PDF document operations.
/// Implements the format-agnostic IOfficeDocumentService interface.
/// </summary>
public interface IPdfDocumentService : IOfficeDocumentService
{
    // PDF-specific operations
    DocumentResult AddWatermark(string filePath, string text, WatermarkOptions? options = null);
    DocumentResult MergeDocuments(string outputPath, params string[] inputPdfs);
    DocumentResult ExtractPages(string filePath, int[] pageNumbers, string outputPath);
    ContentResult GetPageText(string filePath, int pageNumber);
    ContentResult GetPageRange(string filePath, int startPage, int endPage);
    DocumentResult AddEncryption(string filePath, string userPassword, string? ownerPassword = null);
    DocumentResult RemoveEncryption(string filePath, string password);
}

/// <summary>
/// Watermark options for PDF documents.
/// </summary>
public record WatermarkOptions(
    double Opacity = 0.3,
    double Rotation = -45,
    string? FontColor = null,
    int? FontSize = null,
    WatermarkPosition Position = WatermarkPosition.Background
);

public enum WatermarkPosition { Background, Foreground }
