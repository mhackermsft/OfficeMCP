using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeMCP.Tools;

namespace OfficeMCP.Tests;

/// <summary>
/// Tests for the consolidated office_* tools (Tier 3: Advanced Operations).
/// </summary>
[Collection("Office Tests")]
public class ConsolidatedToolsAdvancedTests
{
    private readonly TestFixture _fixture;
    private readonly OfficeDocumentToolsConsolidated _tools;

    public ConsolidatedToolsAdvancedTests(TestFixture fixture)
    {
        _fixture = fixture;
        _tools = fixture.ServiceProvider.GetRequiredService<OfficeDocumentToolsConsolidated>();
    }

    #region office_batch Tests

    [Fact]
    public void BatchOperations_MultipleOperations_ExecutesAll()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("batch_test.docx");
        _tools.CreateDocument(filePath);
        
        var operations = """
        [
            {"type": "heading", "content": "Batch Test Document", "level": 1},
            {"type": "paragraph", "content": "This paragraph was added via batch."},
            {"type": "bulletlist", "content": "Item A\nItem B\nItem C"},
            {"type": "pagebreak"}
        ]
        """;

        // Act
        var result = _tools.BatchOperations(filePath, operationsJson: operations);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.Equal(4, json.RootElement.GetProperty("TotalOperations").GetInt32());
        Assert.Equal(4, json.RootElement.GetProperty("SuccessfulOperations").GetInt32());
        Assert.Equal(0, json.RootElement.GetProperty("FailedOperations").GetInt32());
    }

    [Fact]
    public void BatchOperations_SomeFailures_ReportsPartialSuccess()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("batch_partial.docx");
        _tools.CreateDocument(filePath);
        
        var operations = """
        [
            {"type": "paragraph", "content": "Valid paragraph"},
            {"type": "unknowntype", "content": "This should fail"},
            {"type": "heading", "content": "Valid heading", "level": 2}
        ]
        """;

        // Act
        var result = _tools.BatchOperations(filePath, operationsJson: operations);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.Equal(3, json.RootElement.GetProperty("TotalOperations").GetInt32());
        Assert.True(json.RootElement.GetProperty("FailedOperations").GetInt32() >= 1);
    }

    [Fact]
    public void BatchOperations_InvalidJson_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("batch_invalid.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.BatchOperations(filePath, operationsJson: "not valid json");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
    }

    [Fact]
    public void BatchOperations_EmptyOperations_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("batch_empty.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.BatchOperations(filePath, operationsJson: "[]");
        var json = JsonDocument.Parse(result);




        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
    }

    #endregion

    #region office_merge Tests

    [Fact]
    public void MergeDocuments_MultiplePdfs_CreatesMergedFile()
    {
        // Arrange
        var pdf1 = _fixture.GetTestFilePath("merge_source1.pdf");
        var pdf2 = _fixture.GetTestFilePath("merge_source2.pdf");
        var outputPath = _fixture.GetTestFilePath("merged_output.pdf");
        
        _tools.CreateDocument(pdf1, markdown: "# Document 1\n\nFirst document content.");
        _tools.CreateDocument(pdf2, markdown: "# Document 2\n\nSecond document content.");

        var inputPaths = $"[\"{pdf1.Replace("\\", "\\\\")}\",\"{pdf2.Replace("\\", "\\\\")}\"]";

        // Act
        var result = _tools.MergeDocuments(outputPath, inputPathsJson: inputPaths);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.True(File.Exists(outputPath), "Merged PDF was not created");
    }

    [Fact]
    public void MergeDocuments_UnsupportedFormat_ReturnsError()
    {
        // Arrange
        var outputPath = _fixture.GetTestFilePath("merged.docx");
        var inputPaths = "[\"file1.docx\",\"file2.docx\"]";

        // Act
        var result = _tools.MergeDocuments(outputPath, inputPathsJson: inputPaths);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("not supported", json.RootElement.GetProperty("Message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MergeDocuments_EmptyInputs_ReturnsError()
    {
        // Arrange
        var outputPath = _fixture.GetTestFilePath("merged_empty.pdf");


        // Act
        var result = _tools.MergeDocuments(outputPath, inputPathsJson: "[]");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
    }


    #endregion

    #region office_pdf_pages Tests

    [Fact]
    public void PdfPageOperations_AddWatermark_AddsWatermark()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("watermark_test.pdf");
        _tools.CreateDocument(filePath, markdown: "# Test Document\n\nContent for watermark testing.");

        // Act
        var result = _tools.PdfPageOperations(
            filePath, 
            operation: "watermark",
            watermarkText: "CONFIDENTIAL",
            opacity: 0.3,
            rotation: -45);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void PdfPageOperations_GetPage_ReturnsPageInfo()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("getpage_test.pdf");
        _tools.CreateDocument(filePath, markdown: "# Page Content\n\nSome text here.");

        // Act
        var result = _tools.PdfPageOperations(filePath, operation: "get_page", pageNumbers: "1");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void PdfPageOperations_ExtractPages_CreatesNewPdf()
    {
        // Arrange
        var sourcePath = _fixture.GetTestFilePath("extract_source.pdf");
        var outputPath = _fixture.GetTestFilePath("extracted_pages.pdf");
        _tools.CreateDocument(sourcePath, markdown: "# Page 1\n\nContent.");

        // Act
        var result = _tools.PdfPageOperations(
            sourcePath, 
            operation: "extract_pages",
            pageNumbers: "[1]",
            outputPath: outputPath);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.True(File.Exists(outputPath), "Extracted PDF was not created");
    }

    [Fact]
    public void PdfPageOperations_NonPdfFile_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("not_a_pdf.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.PdfPageOperations(filePath, operation: "watermark", watermarkText: "TEST");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("PDF", json.RootElement.GetProperty("Message").GetString());
    }

    [Fact]
    public void PdfPageOperations_UnknownOperation_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("unknown_op.pdf");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.PdfPageOperations(filePath, operation: "invalid_operation");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
    }

    #endregion
}
