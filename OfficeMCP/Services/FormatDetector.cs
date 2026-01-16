using Microsoft.Extensions.DependencyInjection;

namespace OfficeMCP.Services;

/// <summary>
/// Detects document format from file extension and routes to appropriate service.
/// </summary>
public static class FormatDetector
{
    /// <summary>
    /// Detects format from file extension.
    /// </summary>
    public static string DetectFormat(string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        return extension switch
        {
            ".docx" => "docx",
            ".xlsx" => "xlsx",
            ".pptx" => "pptx",
            ".pdf" => "pdf",
            ".md" => "md",
            _ => throw new InvalidOperationException(
                $"Unsupported format: {extension}. " +
                "Supported: .docx, .xlsx, .pptx, .pdf, .md")
        };
    }

    /// <summary>
    /// Gets the appropriate service for a given format.
    /// Returns IOfficeDocumentService for Word and PDF.
    /// For Excel and PowerPoint, use GetExcelService/GetPowerPointService.
    /// </summary>
    public static IOfficeDocumentService GetService(string format, IServiceProvider services)
    {
        return format.ToLowerInvariant() switch
        {
            "docx" => services.GetRequiredService<IWordDocumentService>(),
            "pdf" => services.GetRequiredService<IPdfDocumentService>(),
            _ => throw new InvalidOperationException(
                $"Format '{format}' does not implement IOfficeDocumentService. " +
                "Use format-specific methods for xlsx and pptx.")
        };
    }

    /// <summary>
    /// Gets the Excel service.
    /// </summary>
    public static IExcelDocumentService GetExcelService(IServiceProvider services)
        => services.GetRequiredService<IExcelDocumentService>();

    /// <summary>
    /// Gets the PowerPoint service.
    /// </summary>
    public static IPowerPointDocumentService GetPowerPointService(IServiceProvider services)
        => services.GetRequiredService<IPowerPointDocumentService>();

    /// <summary>
    /// Gets the PDF service.
    /// </summary>
    public static IPdfDocumentService GetPdfService(IServiceProvider services)
        => services.GetRequiredService<IPdfDocumentService>();

    /// <summary>
    /// Checks if a format is supported.
    /// </summary>
    public static bool IsSupported(string format)
    {
        return format.ToLowerInvariant() switch
        {
            "docx" or "xlsx" or "pptx" or "pdf" or "md" => true,
            _ => false
        };
    }

    /// <summary>
    /// Checks if a format uses the unified IOfficeDocumentService interface.
    /// </summary>
    public static bool UsesUnifiedInterface(string format)
    {
        return format.ToLowerInvariant() switch
        {
            "docx" or "pdf" => true,
            _ => false
        };
    }
}
