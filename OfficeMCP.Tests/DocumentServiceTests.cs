using Microsoft.Extensions.DependencyInjection;
using OfficeMCP.Models;
using OfficeMCP.Services;

namespace OfficeMCP.Tests;

/// <summary>
/// Tests for individual document services to ensure they work correctly.
/// </summary>
[Collection("Office Tests")]
public class DocumentServiceTests
{
    private readonly TestFixture _fixture;

    public DocumentServiceTests(TestFixture fixture)
    {
        _fixture = fixture;
    }

    #region WordDocumentService Tests

    [Fact]
    public void WordService_CreateDocument_CreatesValidDocx()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IWordDocumentService>();
        var filePath = _fixture.GetTestFilePath("word_service_test.docx");

        // Act
        var result = service.CreateDocument(filePath, "Test Title");

        // Assert
        Assert.True(result.Success, result.Message);
        Assert.True(File.Exists(filePath));
    }

    [Fact]
    public void WordService_AddMarkdownContent_ParsesMarkdown()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IWordDocumentService>();
        var filePath = _fixture.GetTestFilePath("word_markdown_test.docx");
        service.CreateDocument(filePath);

        // Act
        var result = service.AddMarkdownContent(filePath, "# Heading\n\n**Bold** and *italic* text.\n\n- Bullet 1\n- Bullet 2");

        // Assert
        Assert.True(result.Success, result.Message);
    }

    [Fact]
    public void WordService_ConvertToMarkdown_ReturnsMarkdown()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IWordDocumentService>();
        var filePath = _fixture.GetTestFilePath("word_to_md_test.docx");
        service.CreateDocument(filePath, "Test");
        service.AddParagraph(filePath, "Test paragraph content");

        // Act
        var result = service.ConvertToMarkdown(filePath);

        // Assert
        Assert.True(result.Success, result.ErrorMessage ?? "Failed");
        Assert.NotNull(result.Content);
    }

    #endregion

    #region ExcelDocumentService Tests

    [Fact]
    public void ExcelService_CreateWorkbook_CreatesValidXlsx()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IExcelDocumentService>();
        var filePath = _fixture.GetTestFilePath("excel_service_test.xlsx");

        // Act
        var result = service.CreateWorkbook(filePath, "TestSheet");

        // Assert
        Assert.True(result.Success, result.Message);
        Assert.True(File.Exists(filePath));
    }

    [Fact]
    public void ExcelService_SetCellValue_SetsValue()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IExcelDocumentService>();
        var filePath = _fixture.GetTestFilePath("excel_cell_test.xlsx");
        service.CreateWorkbook(filePath, "Sheet1");

        // Act
        var result = service.SetCellValue(filePath, "Sheet1", "A1", "Test Value");

        // Assert
        Assert.True(result.Success, result.Message);
    }

    [Fact]
    public void ExcelService_GetCellValue_ReturnsValue()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IExcelDocumentService>();
        var filePath = _fixture.GetTestFilePath("excel_get_cell_test.xlsx");
        service.CreateWorkbook(filePath, "Sheet1");
        service.SetCellValue(filePath, "Sheet1", "B2", "Expected Value");

        // Act
        var result = service.GetCellValue(filePath, "Sheet1", "B2");

        // Assert
        Assert.True(result.Success, result.ErrorMessage ?? "Failed");
        Assert.Equal("Expected Value", result.Content);
    }

    [Fact]
    public void ExcelService_AddTable_CreatesTable()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IExcelDocumentService>();
        var filePath = _fixture.GetTestFilePath("excel_table_test.xlsx");
        service.CreateWorkbook(filePath, "Sheet1");
        var data = new[] { new[] { "Name", "Age" }, new[] { "Alice", "30" }, new[] { "Bob", "25" } };

        // Act
        var result = service.AddTable(filePath, "Sheet1", "A1", data, hasHeaders: true);

        // Assert
        Assert.True(result.Success, result.Message);
    }

    #endregion

    #region PowerPointDocumentService Tests

    [Fact]
    public void PowerPointService_CreatePresentation_CreatesValidPptx()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IPowerPointDocumentService>();
        var filePath = _fixture.GetTestFilePath("pptx_service_test.pptx");

        // Act
        var result = service.CreatePresentation(filePath, "Test Presentation");

        // Assert
        Assert.True(result.Success, result.Message);
        Assert.True(File.Exists(filePath));
    }

    [Fact]
    public void PowerPointService_AddSlide_AddsSlide()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IPowerPointDocumentService>();
        var filePath = _fixture.GetTestFilePath("pptx_slide_test.pptx");
        service.CreatePresentation(filePath);

        // Act
        var result = service.AddSlide(filePath);

        // Assert
        Assert.True(result.Success, result.Message);
    }

    [Fact]
    public void PowerPointService_GetSlideCount_ReturnsCount()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IPowerPointDocumentService>();
        var filePath = _fixture.GetTestFilePath("pptx_count_test.pptx");
        service.CreatePresentation(filePath, "Title");

        // Act
        var result = service.GetSlideCount(filePath);


        // Assert
        Assert.True(result.Success, result.ErrorMessage ?? "Failed");
        Assert.NotNull(result.Content);
    }


    #endregion

    #region PdfDocumentService Tests

    [Fact]
    public void PdfService_CreateDocument_CreatesValidPdf()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IPdfDocumentService>();
        var filePath = _fixture.GetTestFilePath("pdf_service_test.pdf");

        // Act
        var result = service.CreateDocument(filePath, "Test PDF");

        // Assert
        Assert.True(result.Success, result.Message);
        Assert.True(File.Exists(filePath));
    }

    [Fact]
    public void PdfService_AddMarkdownContent_AddsContent()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IPdfDocumentService>();
        var filePath = _fixture.GetTestFilePath("pdf_markdown_test.pdf");
        service.CreateDocument(filePath);

        // Act
        var result = service.AddMarkdownContent(filePath, "# PDF Heading\n\nParagraph text here.\n\n- Bullet point");

        // Assert
        Assert.True(result.Success, result.Message);
    }

    [Fact]
    public void PdfService_AddTable_AddsTable()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IPdfDocumentService>();
        var filePath = _fixture.GetTestFilePath("pdf_table_test.pdf");
        service.CreateDocument(filePath);
        var data = new[] { new[] { "Col1", "Col2" }, new[] { "A", "B" } };

        // Act
        var result = service.AddTable(filePath, data);

        // Assert
        Assert.True(result.Success, result.Message);
    }

    [Fact]
    public void PdfService_AddWatermark_AddsWatermark()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IPdfDocumentService>();
        var filePath = _fixture.GetTestFilePath("pdf_watermark_test.pdf");
        service.CreateDocument(filePath);
        service.AddMarkdownContent(filePath, "# Test\n\nContent");

        // Act
        var result = service.AddWatermark(filePath, "DRAFT");

        // Assert
        Assert.True(result.Success, result.Message);
    }

    [Fact]
    public void PdfService_MergeDocuments_MergesPdfs()
    {
        // Arrange
        var service = _fixture.ServiceProvider.GetRequiredService<IPdfDocumentService>();
        var pdf1 = _fixture.GetTestFilePath("pdf_merge1.pdf");
        var pdf2 = _fixture.GetTestFilePath("pdf_merge2.pdf");
        var output = _fixture.GetTestFilePath("pdf_merged.pdf");
        
        service.CreateDocument(pdf1);
        service.AddMarkdownContent(pdf1, "# Doc 1");
        service.CreateDocument(pdf2);
        service.AddMarkdownContent(pdf2, "# Doc 2");

        // Act
        var result = service.MergeDocuments(output, pdf1, pdf2);

        // Assert
        Assert.True(result.Success, result.Message);
        Assert.True(File.Exists(output));
    }

    #endregion
}
