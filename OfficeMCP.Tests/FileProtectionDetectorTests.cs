using Microsoft.Extensions.DependencyInjection;
using OfficeMCP.Services;

namespace OfficeMCP.Tests;

/// <summary>
/// Tests for the OfficeFileProtectionDetector utility class.
/// </summary>
[Collection("Office Tests")]
public class FileProtectionDetectorTests
{
    private readonly TestFixture _fixture;

    public FileProtectionDetectorTests(TestFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void CheckFileProtection_NonExistentFile_ReturnsNotProtected()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("nonexistent_protection_test.docx");

        // Act
        var result = OfficeFileProtectionDetector.CheckFileProtection(filePath);

        // Assert
        Assert.False(result.IsProtected);
        Assert.False(result.IsValidOfficeFormat);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void CheckFileProtection_ValidDocx_ReturnsNotProtected()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("valid_docx_protection.docx");
        var wordService = _fixture.ServiceProvider.GetRequiredService<IWordDocumentService>();
        wordService.CreateDocument(filePath);

        // Act
        var result = OfficeFileProtectionDetector.CheckFileProtection(filePath);

        // Assert
        Assert.False(result.IsEncrypted);
        Assert.True(result.IsValidOfficeFormat);
        Assert.Equal("None", result.ProtectionType);
    }

    [Fact]
    public void CheckFileProtection_ValidXlsx_ReturnsNotProtected()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("valid_xlsx_protection.xlsx");
        var excelService = _fixture.ServiceProvider.GetRequiredService<IExcelDocumentService>();
        excelService.CreateWorkbook(filePath);

        // Act
        var result = OfficeFileProtectionDetector.CheckFileProtection(filePath);

        // Assert
        Assert.False(result.IsEncrypted);
        Assert.True(result.IsValidOfficeFormat);
    }

    [Fact]
    public void CheckFileProtection_ValidPptx_ReturnsNotProtected()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("valid_pptx_protection.pptx");
        var pptService = _fixture.ServiceProvider.GetRequiredService<IPowerPointDocumentService>();
        pptService.CreatePresentation(filePath);

        // Act
        var result = OfficeFileProtectionDetector.CheckFileProtection(filePath);

        // Assert
        Assert.False(result.IsEncrypted);
        Assert.True(result.IsValidOfficeFormat);
    }

    [Fact]
    public void CheckFileProtection_ValidPdf_ReturnsNotProtected()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("valid_pdf_protection.pdf");
        var pdfService = _fixture.ServiceProvider.GetRequiredService<IPdfDocumentService>();
        pdfService.CreateDocument(filePath);

        // Act
        var result = OfficeFileProtectionDetector.CheckFileProtection(filePath);

        // Assert
        Assert.False(result.IsEncrypted);
        Assert.True(result.IsValidOfficeFormat);
    }

    [Fact]
    public void CheckFileProtection_UnsupportedFormat_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("unsupported.txt");
        File.WriteAllText(filePath, "test content");

        // Act
        var result = OfficeFileProtectionDetector.CheckFileProtection(filePath);

        // Assert
        Assert.False(result.IsValidOfficeFormat);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void GetProtectionErrorMessage_UnprotectedFile_ReturnsEmpty()
    {
        // Arrange
        var info = new OfficeFileProtectionDetector.FileProtectionInfo(
            IsProtected: false,
            IsEncrypted: false,
            MayHaveSensitivityLabel: false,
            IsValidOfficeFormat: true,
            ProtectionType: "None"
        );

        // Act
        var message = OfficeFileProtectionDetector.GetProtectionErrorMessage(info, "test.docx");

        // Assert
        Assert.Empty(message);
    }

    [Fact]
    public void GetProtectionErrorMessage_EncryptedFile_ReturnsMessage()
    {
        // Arrange
        var info = new OfficeFileProtectionDetector.FileProtectionInfo(
            IsProtected: true,
            IsEncrypted: true,
            MayHaveSensitivityLabel: false,
            IsValidOfficeFormat: false,
            ProtectionType: "Encrypted"
        );

        // Act
        var message = OfficeFileProtectionDetector.GetProtectionErrorMessage(info, "protected.docx");

        // Assert
        Assert.NotEmpty(message);
        Assert.Contains("protected.docx", message);
        Assert.Contains("encrypted", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetProtectionErrorMessage_SensitivityLabel_ReturnsMessage()
    {
        // Arrange
        var info = new OfficeFileProtectionDetector.FileProtectionInfo(
            IsProtected: true,
            IsEncrypted: true,
            MayHaveSensitivityLabel: true,
            IsValidOfficeFormat: false,
            ProtectionType: "Encrypted"
        );

        // Act
        var message = OfficeFileProtectionDetector.GetProtectionErrorMessage(info, "labeled.docx");

        // Assert
        Assert.NotEmpty(message);
        Assert.Contains("sensitivity label", message, StringComparison.OrdinalIgnoreCase);
    }
}
