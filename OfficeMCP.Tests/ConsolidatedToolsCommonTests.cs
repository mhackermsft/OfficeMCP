using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeMCP.Tools;

namespace OfficeMCP.Tests;

/// <summary>
/// Tests for the consolidated office_* tools (Tier 2: Common Operations).
/// </summary>
[Collection("Office Tests")]
public class ConsolidatedToolsCommonTests
{
    private readonly TestFixture _fixture;
    private readonly OfficeDocumentToolsConsolidated _tools;

    public ConsolidatedToolsCommonTests(TestFixture fixture)
    {
        _fixture = fixture;
        _tools = fixture.ServiceProvider.GetRequiredService<OfficeDocumentToolsConsolidated>();
    }

    #region office_add_element Tests

    [Fact]
    public void AddElement_Paragraph_AddsToDocument()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("element_paragraph.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.AddElement(filePath, elementType: "paragraph", content: "This is a test paragraph.");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddElement_Heading_AddsWithLevel()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("element_heading.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.AddElement(filePath, elementType: "heading", content: "Section Title", level: 2);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddElement_Table_AddsTableFromJson()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("element_table.docx");
        _tools.CreateDocument(filePath);
        var tableData = "[[\"Col1\",\"Col2\"],[\"A\",\"B\"],[\"C\",\"D\"]]";

        // Act
        var result = _tools.AddElement(filePath, elementType: "table", content: tableData);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddElement_PageBreak_AddsBreak()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("element_pagebreak.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.AddElement(filePath, elementType: "pagebreak");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddElement_BulletList_AddsList()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("element_bullets.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.AddElement(filePath, elementType: "bulletlist", content: "Item 1\nItem 2\nItem 3");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddElement_NumberedList_AddsList()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("element_numbered.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.AddElement(filePath, elementType: "numberedlist", content: "First\nSecond\nThird");
        var json = JsonDocument.Parse(result);


        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddElement_Pdf_AddsTable()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("element_table.pdf");
        _tools.CreateDocument(filePath);
        var tableData = "[[\"Header1\",\"Header2\"],[\"Row1Col1\",\"Row1Col2\"]]";

        // Act
        var result = _tools.AddElement(filePath, elementType: "table", content: tableData);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddElement_NonExistentFile_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("nonexistent_element.docx");

        // Act
        var result = _tools.AddElement(filePath, elementType: "paragraph", content: "Test");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
    }

    #endregion

    #region office_add_header_footer Tests

    [Fact]
    public void AddHeaderFooter_Header_AddsToWord()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("header_test.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.AddHeaderFooter(
            filePath, 
            location: "header",
            leftContent: "Company Name",
            centerContent: "Document Title",
            rightContent: "Confidential");
        var json = JsonDocument.Parse(result);


        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddHeaderFooter_Footer_WithPageNumber()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("footer_test.docx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.AddHeaderFooter(
            filePath, 
            location: "footer",
            centerContent: "Page",
            includePageNumber: true);
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddHeaderFooter_Pdf_AddsHeader()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("header_test.pdf");
        _tools.CreateDocument(filePath, markdown: "# Page 1\n\nContent here.");

        // Act
        var result = _tools.AddHeaderFooter(filePath, location: "header", centerContent: "PDF Header");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void AddHeaderFooter_UnsupportedFormat_ReturnsError()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("header_test.xlsx");
        _tools.CreateDocument(filePath);

        // Act
        var result = _tools.AddHeaderFooter(filePath, location: "header", centerContent: "Test");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.False(json.RootElement.GetProperty("Success").GetBoolean());
    }

    #endregion

    #region office_extract Tests

    [Fact]
    public void ExtractContent_Text_ReturnsText()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("extract_text.docx");
        _tools.CreateDocument(filePath, markdown: "# Title\n\nSome content to extract.");

        // Act
        var result = _tools.ExtractContent(filePath, extractType: "text");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    [Fact]
    public void ExtractContent_Metadata_ReturnsMetadata()
    {
        // Arrange
        var filePath = _fixture.GetTestFilePath("extract_metadata.docx");
        _tools.CreateDocument(filePath, title: "Extract Test");

        // Act
        var result = _tools.ExtractContent(filePath, extractType: "metadata");
        var json = JsonDocument.Parse(result);

        // Assert
        Assert.True(json.RootElement.GetProperty("Success").GetBoolean(), $"Failed: {result}");
    }

    #endregion
}
