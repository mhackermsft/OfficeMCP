using OfficeMCP.Services;

namespace OfficeMCP.Tests;

/// <summary>
/// Tests for the FormatDetector utility class.
/// </summary>
public class FormatDetectorTests
{
    [Theory]
    [InlineData("document.docx", "docx")]
    [InlineData("workbook.xlsx", "xlsx")]
    [InlineData("presentation.pptx", "pptx")]
    [InlineData("document.pdf", "pdf")]
    [InlineData("readme.md", "md")]
    [InlineData("C:\\path\\to\\file.docx", "docx")]
    [InlineData("/unix/path/file.xlsx", "xlsx")]
    [InlineData("file.DOCX", "docx")]  // Case insensitive
    [InlineData("file.PDF", "pdf")]    // Case insensitive
    public void DetectFormat_ValidExtensions_ReturnsCorrectFormat(string filePath, string expectedFormat)
    {
        // Act
        var result = FormatDetector.DetectFormat(filePath);

        // Assert
        Assert.Equal(expectedFormat, result);
    }

    [Theory]
    [InlineData("file.txt")]
    [InlineData("file.doc")]   // Old Word format not supported
    [InlineData("file.xls")]   // Old Excel format not supported
    [InlineData("file.ppt")]   // Old PowerPoint format not supported
    [InlineData("file.jpg")]
    [InlineData("file")]       // No extension
    public void DetectFormat_InvalidExtensions_ThrowsException(string filePath)
    {
        // Act & Assert
        Assert.Throws<InvalidOperationException>(() => FormatDetector.DetectFormat(filePath));
    }

    [Theory]
    [InlineData("docx", true)]
    [InlineData("xlsx", true)]
    [InlineData("pptx", true)]
    [InlineData("pdf", true)]
    [InlineData("md", true)]
    [InlineData("txt", false)]
    [InlineData("doc", false)]
    [InlineData("", false)]
    public void IsSupported_ReturnsCorrectResult(string format, bool expectedSupported)
    {
        // Act
        var result = FormatDetector.IsSupported(format);

        // Assert
        Assert.Equal(expectedSupported, result);
    }

    [Theory]
    [InlineData("docx", true)]
    [InlineData("pdf", true)]
    [InlineData("xlsx", false)]
    [InlineData("pptx", false)]
    public void UsesUnifiedInterface_ReturnsCorrectResult(string format, bool expectedUnified)
    {
        // Act
        var result = FormatDetector.UsesUnifiedInterface(format);

        // Assert
        Assert.Equal(expectedUnified, result);
    }
}
