using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeMCP.Tools;

namespace OfficeMCP.Tests;

/// <summary>
/// Tests for the consolidated office_* tools (Tier 1: Core Operations).
/// </summary>
[Collection("Office Tests")]
public class ConsolidatedToolsCoreTests
{
    private readonly TestFixture _fixture;
    private readonly OfficeDocumentToolsConsolidated _tools;

    public ConsolidatedToolsCoreTests(TestFixture fixture)
    {
        _fixture = fixture;
        _tools = fixture.ServiceProvider.GetRequiredService<OfficeDocumentToolsConsolidated>();
    }

    #region office_create Tests

    [Theory]
    [InlineData("test.docx", "docx")]
    [InlineData("test.xlsx", "xlsx")]
    [InlineData("test.pptx", "pptx")]
    [InlineData("test.pdf", "pdf")]
    public void CreateDocument_AllFormats_CreatesFile(string fileName, string expectedFormat)
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath(fileName);

        // Act
        var result = _tools.CreateDocument(filePath, title: "Test Document");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.Equal(expectedFormat, json.RootElement.GetProperty("Format").GetString());
        Assert.True(File.Exists(filePath), $"File was not created: {filePath}");
    }

    [Fact]
    public void CreateDocument_WithMarkdown_AddsContent()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("markdown_test.docx");
        var markdown = "# Heading 1\n\nThis is a paragraph.\n\n- Item 1\n- Item 2";

        // Act
        var result = _tools.CreateDocument(filePath, title: "Markdown Test", markdown: markdown);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.True(File.Exists(filePath));
    }

    [Fact]
    public void CreateDocument_Excel_WithJsonData_CreatesTable()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("data_test.xlsx");
        var jsonData = "[[\"Name\",\"Age\"],[\"Alice\",\"30\"],[\"Bob\",\"25\"]]";

        // Act
        var result = _tools.CreateDocument(filePath, title: "Sheet1", markdown: jsonData);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.True(File.Exists(filePath));
    }

    [Fact]
    public void CreateDocument_InvalidPath_ReturnsError()
    {
        // Arrange
        var filePath = "Z:\\NonExistent\\Path\\test.docx";

        // Act
        var result = _tools.CreateDocument(filePath);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
    }

    [Fact]
    public void CreateDocument_UnsupportedFormat_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("test.xyz");

        // Act
        var result = _tools.CreateDocument(filePath);
        var json = JsonDocument.Parse(result);
        
        // Assert - should return error result, not throw exception
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("unsupported", json.RootElement.GetProperty("Message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region office_read Tests

    [Theory]
    [InlineData("read_test.docx")]
    [InlineData("read_test.pdf")]
    public void ReadDocument_ExistingFile_ReturnsContent(string fileName)
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath(fileName);
        _tools.CreateDocument(filePath, markdown: "# Test Content\n\nHello World");

        // Act
        var result = _tools.ReadDocument(filePath);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void ReadDocument_NonExistentFile_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("nonexistent.docx");

        // Act
        var result = _tools.ReadDocument(filePath);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("not found", json.RootElement.GetProperty("Message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReadDocument_Excel_ReadsAllSheets()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("excel_read_test.xlsx");
        _tools.CreateDocument(filePath, title: "TestSheet");

        // Act
        var result = _tools.ReadDocument(filePath, readType: "all");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.Equal("xlsx", json.RootElement.GetProperty("Format").GetString());
    }

    [Fact]
    public void ReadDocument_PowerPoint_ReadsAllSlides()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("pptx_read_test.pptx");
        _tools.CreateDocument(filePath, title: "Test Presentation");

        // Act
        var result = _tools.ReadDocument(filePath, readType: "all");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.Equal("pptx", json.RootElement.GetProperty("Format").GetString());
    }

    #endregion

    #region office_write Tests

    [Fact]
    public void WriteDocument_Word_AppendsContent()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("write_test.docx");
        _tools.CreateDocument(filePath, title: "Initial");

        // Act
        var result = _tools.WriteDocument(filePath, content: "## New Section\n\nAdditional content here.");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void WriteDocument_Pdf_AppendsContent()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("write_test.pdf");
        _tools.CreateDocument(filePath, title: "Initial");

        // Act
        var result = _tools.WriteDocument(filePath, content: "# Added Heading\n\nMore text.");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void WriteDocument_NonExistentFile_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("nonexistent_write.docx");

        // Act
        var result = _tools.WriteDocument(filePath, content: "Some content");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
    }

    [Fact]
    public void WriteDocument_EmptyContent_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("empty_content.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.WriteDocument(filePath, content: "");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
    }

    #endregion

    #region office_metadata Tests

    [Theory]
    [InlineData("metadata_test.docx", "docx")]
    [InlineData("metadata_test.xlsx", "xlsx")]
    [InlineData("metadata_test.pptx", "pptx")]
    [InlineData("metadata_test.pdf", "pdf")]
    public void GetMetadata_AllFormats_ReturnsMetadata(string fileName, string expectedFormat)
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath(fileName);
        _tools.CreateDocument(filePath, title: "Test");

        // Act
        var result = _tools.GetMetadata(filePath);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.Equal(expectedFormat, json.RootElement.GetProperty("Format").GetString());
        Assert.True(json.RootElement.TryGetProperty("FileSize", out _));
        Assert.True(json.RootElement.TryGetProperty("CreatedDate", out _));
        Assert.True(json.RootElement.TryGetProperty("ModifiedDate", out _));
    }

    [Fact]
    public void GetMetadata_WithStructure_ReturnsDetailedInfo()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("metadata_structure.xlsx");
        _tools.CreateDocument(filePath, title: "Sheet1");

        // Act
        var result = _tools.GetMetadata(filePath, includeStructure: true);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    #endregion

    #region office_convert Tests

    [Fact]
    public void ConvertDocument_WordToMarkdown_CreatesMarkdownFile()
    {
        // Arrange
        var sourcePath = _fixture.GetTestFilePath("convert_source.docx");
        var outputPath = _fixture.GetTestFilePath("convert_output.md");
        _tools.CreateDocument(sourcePath, markdown: "# Test\n\nParagraph content.");

        // Act
        var result = _tools.ConvertDocument(sourcePath, targetFormat: "md", outputPath: outputPath);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
        Assert.True(File.Exists(outputPath), "Markdown file was not created");
    }

    [Fact]
    public void ConvertDocument_UnsupportedConversion_ReturnsError()
    {
        // Arrange
        var sourcePath = _fixture.GetTestFilePath("convert_unsupported.xlsx");
        _tools.CreateDocument(sourcePath);

        // Act
        var result = _tools.ConvertDocument(sourcePath, targetFormat: "docx");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
        Assert.Contains("not yet implemented", json.RootElement.GetProperty("Message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
