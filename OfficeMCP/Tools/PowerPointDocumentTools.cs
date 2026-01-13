using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// MCP Tools for creating and manipulating PowerPoint presentations.
/// </summary>
[McpServerToolType]
public sealed class PowerPointDocumentTools(IPowerPointDocumentService powerPointService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    [McpServerTool, Description("Create a new PowerPoint presentation (.pptx). Optionally add a title slide.")]
    public string CreatePowerPointPresentation(
        [Description("Full file path for the new presentation (e.g., C:/Documents/slides.pptx)")] string filePath,
        [Description("Optional title for the first slide")] string? title = null)
    {
        var result = powerPointService.CreatePresentation(filePath, title);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a new slide to the presentation.")]
    public string AddSlideToPowerPoint(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Background color as hex (e.g., 'FFFFFF' for white)")] string? backgroundColor = null)
    {
        var layoutOptions = new SlideLayoutOptions(BackgroundColor: backgroundColor);
        var result = powerPointService.AddSlide(filePath, layoutOptions);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a title and optional subtitle to a slide.")]
    public string AddTitleToPowerPoint(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("The title text")] string title,
        [Description("Optional subtitle text")] string? subtitle = null,
        [Description("Make text bold")] bool bold = true,
        [Description("Font size in points (default: 44 for title)")] int fontSize = 44,
        [Description("Font color as hex (e.g., '000000' for black)")] string? fontColor = null)
    {
        var textFormat = new TextFormatting(Bold: bold, FontSize: fontSize, FontColor: fontColor);
        var result = powerPointService.AddTitle(filePath, slideIndex, title, subtitle, textFormat);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a text box to a slide at a specific position.")]
    public string AddTextBoxToPowerPoint(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("The text content")] string text,
        [Description("X position in inches from left edge")] double xInches = 1.0,
        [Description("Y position in inches from top edge")] double yInches = 2.0,
        [Description("Width in inches")] double widthInches = 8.0,
        [Description("Height in inches")] double heightInches = 1.0,
        [Description("Make text bold")] bool bold = false,
        [Description("Font size in points")] int fontSize = 18,
        [Description("Font color as hex")] string? fontColor = null,
        [Description("Background color as hex (leave empty for transparent)")] string? backgroundColor = null,
        [Description("Border color as hex (leave empty for no border)")] string? borderColor = null)
    {
        var textFormat = new TextFormatting(Bold: bold, FontSize: fontSize, FontColor: fontColor);
        var options = new TextBoxOptions(
            X: (long)(xInches * 914400),
            Y: (long)(yInches * 914400),
            Width: (long)(widthInches * 914400),
            Height: (long)(heightInches * 914400),
            BackgroundColor: backgroundColor,
            BorderColor: borderColor,
            TextFormat: textFormat
        );
        
        var result = powerPointService.AddTextBox(filePath, slideIndex, text, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add bullet points to a slide.")]
    public string AddBulletPointsToPowerPoint(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Bullet points as JSON array, e.g., [\"Point 1\", \"Point 2\", \"Point 3\"]")] string pointsJson,
        [Description("X position in inches from left edge")] double xInches = 1.0,
        [Description("Y position in inches from top edge")] double yInches = 2.0,
        [Description("Width in inches")] double widthInches = 8.0,
        [Description("Height in inches")] double heightInches = 4.0,
        [Description("Font size in points")] int fontSize = 18,
        [Description("Font color as hex")] string? fontColor = null)
    {
        try
        {
            var points = JsonSerializer.Deserialize<string[]>(pointsJson, JsonOptions);
            if (points == null || points.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty points array"), JsonOptions);
            }

            var textFormat = new TextFormatting(FontSize: fontSize, FontColor: fontColor);
            var options = new TextBoxOptions(
                X: (long)(xInches * 914400),
                Y: (long)(yInches * 914400),
                Width: (long)(widthInches * 914400),
                Height: (long)(heightInches * 914400),
                TextFormat: textFormat
            );
            
            var result = powerPointService.AddBulletPoints(filePath, slideIndex, points, options);
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description("Add an image to a slide at a specific position.")]
    public string AddImageToPowerPoint(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Full path to the image file")] string imagePath,
        [Description("X position in inches from left edge")] double xInches = 1.0,
        [Description("Y position in inches from top edge")] double yInches = 2.0,
        [Description("Image width in inches")] double widthInches = 4.0,
        [Description("Image height in inches")] double heightInches = 3.0,
        [Description("Alt text for accessibility")] string? altText = null)
    {
        var options = new ImageOptions(
            WidthEmu: (long)(widthInches * 914400),
            HeightEmu: (long)(heightInches * 914400),
            AltText: altText
        );
        
        var result = powerPointService.AddImage(filePath, slideIndex, imagePath, (long)(xInches * 914400), (long)(yInches * 914400), options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a shape to a slide. Shapes include rectangle, ellipse, triangle, arrow, star, heart, and more.")]
    public string AddShapeToPowerPoint(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Shape type: rectangle, roundRectangle, ellipse, triangle, diamond, pentagon, hexagon, arrow, leftArrow, star, star4, star5, heart")] string shapeType,
        [Description("X position in inches from left edge")] double xInches = 1.0,
        [Description("Y position in inches from top edge")] double yInches = 2.0,
        [Description("Width in inches")] double widthInches = 2.0,
        [Description("Height in inches")] double heightInches = 2.0,
        [Description("Fill color as hex (e.g., '4472C4' for blue)")] string? fillColor = null,
        [Description("Border color as hex")] string? borderColor = null,
        [Description("Border width in points")] double borderWidth = 1.0)
    {
        var options = new ShapeOptions(
            ShapeType: shapeType,
            X: (long)(xInches * 914400),
            Y: (long)(yInches * 914400),
            Width: (long)(widthInches * 914400),
            Height: (long)(heightInches * 914400),
            FillColor: fillColor,
            BorderColor: borderColor,
            BorderWidth: borderWidth
        );
        
        var result = powerPointService.AddShape(filePath, slideIndex, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a table to a slide.")]
    public string AddTableToPowerPoint(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Table data as JSON 2D array where first row can be headers")] string tableDataJson,
        [Description("X position in inches from left edge")] double xInches = 1.0,
        [Description("Y position in inches from top edge")] double yInches = 2.0,
        [Description("Table width in inches")] double widthInches = 8.0,
        [Description("Table height in inches")] double heightInches = 3.0)
    {
        try
        {
            var tableData = JsonSerializer.Deserialize<string[][]>(tableDataJson, JsonOptions);
            if (tableData == null || tableData.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty table data"), JsonOptions);
            }

            var result = powerPointService.AddTable(
                filePath, slideIndex, tableData,
                (long)(xInches * 914400),
                (long)(yInches * 914400),
                (long)(widthInches * 914400),
                (long)(heightInches * 914400)
            );
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description("Set the background color of a slide.")]
    public string SetPowerPointSlideBackground(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Background color as hex (e.g., 'FFFFFF' for white, '000000' for black)")] string color)
    {
        var result = powerPointService.SetSlideBackground(filePath, slideIndex, color);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Delete a slide from the presentation.")]
    public string DeletePowerPointSlide(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Index of the slide to delete (0-based)")] int slideIndex)
    {
        var result = powerPointService.DeleteSlide(filePath, slideIndex);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Duplicate an existing slide.")]
    public string DuplicatePowerPointSlide(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Index of the slide to duplicate (0-based)")] int sourceIndex)
    {
        var result = powerPointService.DuplicateSlide(filePath, sourceIndex);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Move a slide from one position to another.")]
    public string ReorderPowerPointSlide(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Current index of the slide (0-based)")] int fromIndex,
        [Description("New index for the slide (0-based)")] int toIndex)
    {
        var result = powerPointService.ReorderSlide(filePath, fromIndex, toIndex);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Get all text content from a specific slide.")]
    public string GetPowerPointSlideText(
        [Description("Path to the PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex)
    {
        var result = powerPointService.GetSlideText(filePath, slideIndex);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Get all text content from all slides in the presentation.")]
    public string GetPowerPointAllSlidesText(
        [Description("Path to the PowerPoint presentation")] string filePath)
    {
        var result = powerPointService.GetAllSlidesText(filePath);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Get the total number of slides in the presentation.")]
    public string GetPowerPointSlideCount(
        [Description("Path to the PowerPoint presentation")] string filePath)
    {
        var result = powerPointService.GetSlideCount(filePath);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add speaker notes to a slide.")]
    public string AddSpeakerNotesToPowerPoint(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Speaker notes text")] string notes)
    {
        var result = powerPointService.AddSpeakerNotes(filePath, slideIndex, notes);
        return JsonSerializer.Serialize(result, JsonOptions);
    }
}
