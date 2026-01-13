using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// AI-Optimized MCP Tools for PowerPoint presentations. Reduces tool calls through batch operations.
/// </summary>
[McpServerToolType]
public sealed class PowerPointDocumentToolsOptimized(IPowerPointDocumentService powerPointService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    [McpServerTool, Description(@"Create a new PowerPoint presentation and optionally add slides with content in a single call.

**Examples**:
- Simple: {""filePath"": ""C:/presentations/demo.pptx""}
- With title: {""filePath"": ""C:/presentations/demo.pptx"", ""title"": ""Quarterly Review""}
- With title slide content: {""filePath"": ""C:/presentations/demo.pptx"", ""title"": ""Q4 Results"", ""subtitle"": ""Financial Overview""}")]
    public string CreatePowerPointPresentation(
        [Description("Full file path for the new presentation (e.g., C:/Documents/slides.pptx)")] string filePath,
        [Description("Title for the first slide")] string? title = null,
        [Description("Subtitle for the title slide")] string? subtitle = null)
    {
        var result = powerPointService.CreatePresentation(filePath, title);

        if (!result.Success)
        {
            return JsonSerializer.Serialize(result, JsonOptions);
        }

        // If subtitle is provided and we have a title, add it
        if (!string.IsNullOrWhiteSpace(subtitle) && !string.IsNullOrWhiteSpace(title))
        {
            // The CreatePresentation already adds a title slide with the title, we need to add the subtitle
            var textFormat = new TextFormatting(FontSize: 24);
            var subtitleOptions = new TextBoxOptions(
                X: (long)(1.0 * 914400),
                Y: (long)(4.0 * 914400),
                Width: (long)(8.0 * 914400),
                Height: (long)(1.0 * 914400),
                TextFormat: textFormat
            );
            powerPointService.AddTextBox(filePath, 0, subtitle, subtitleOptions);
        }

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Add a new slide to a PowerPoint presentation.")]
    public string AddPowerPointSlide(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Background color as hex (e.g., 'FFFFFF' for white)")] string? backgroundColor = null)
    {
        var layoutOptions = new SlideLayoutOptions(BackgroundColor: backgroundColor);
        var result = powerPointService.AddSlide(filePath, layoutOptions);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Add a title and optional subtitle to a PowerPoint slide.")]
    public string AddPowerPointTitle(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("The title text")] string title,
        [Description("Optional subtitle text")] string? subtitle = null,
        [Description("Make text bold")] bool bold = true,
        [Description("Font size in points (default: 44)")] int fontSize = 44,
        [Description("Font color as hex (e.g., '000000' for black)")] string? fontColor = null)
    {
        var textFormat = new TextFormatting(Bold: bold, FontSize: fontSize, FontColor: fontColor);
        var result = powerPointService.AddTitle(filePath, slideIndex, title, subtitle, textFormat);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Add text content (text box or bullet points) to a PowerPoint slide.")]
    public string AddPowerPointText(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("The text content (for textbox) or first bullet point")] string text,
        [Description("Additional bullet points as JSON array (e.g., [\"Point 2\", \"Point 3\"]) - if provided, creates bullet list")] string? additionalPointsJson = null,
        [Description("X position in inches from left edge")] double xInches = 1.0,
        [Description("Y position in inches from top edge")] double yInches = 2.0,
        [Description("Width in inches")] double widthInches = 8.0,
        [Description("Height in inches")] double heightInches = 1.0,
        [Description("Font size in points")] int fontSize = 18,
        [Description("Font color as hex")] string? fontColor = null)
    {
        var textFormat = new TextFormatting(FontSize: fontSize, FontColor: fontColor);
        var options = new TextBoxOptions(
            X: (long)(xInches * 914400),
            Y: (long)(yInches * 914400),
            Width: (long)(widthInches * 914400),
            Height: (long)(heightInches * 914400),
            TextFormat: textFormat
        );

        DocumentResult result;
        if (!string.IsNullOrWhiteSpace(additionalPointsJson))
        {
            try
            {
                var additionalPoints = JsonSerializer.Deserialize<string[]>(additionalPointsJson, JsonOptions) ?? [];
                var allPoints = new[] { text }.Concat(additionalPoints).ToArray();
                result = powerPointService.AddBulletPoints(filePath, slideIndex, allPoints, options);
            }
            catch (JsonException ex)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON for additional points: {ex.Message}"), JsonOptions);
            }
        }
        else
        {
            result = powerPointService.AddTextBox(filePath, slideIndex, text, options);
        }
        
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Add an image to a PowerPoint slide.")]
    public string AddPowerPointImage(
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

    [McpServerTool, Description(@"Add a table to a PowerPoint slide.")]
    public string AddPowerPointTable(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Table data as JSON 2D array (e.g., [[\"Header1\",\"Header2\"],[\"Row1Col1\",\"Row1Col2\"]])")] string tableDataJson,
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

    [McpServerTool, Description(@"Add speaker notes to a PowerPoint slide.")]
    public string AddPowerPointSpeakerNotes(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Speaker notes text")] string notes)
    {
        var result = powerPointService.AddSpeakerNotes(filePath, slideIndex, notes);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Manage slides (delete, duplicate, reorder) in a PowerPoint presentation.")]
    public string ManagePowerPointSlide(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("Operation: 'delete', 'duplicate', or 'reorder'")] string operation,
        [Description("Slide index for delete/duplicate, or source index for reorder (0-based)")] int slideIndex,
        [Description("Target index for reorder operation (0-based)")] int? targetIndex = null)
    {
        DocumentResult result = operation.ToLowerInvariant() switch
        {
            "delete" => powerPointService.DeleteSlide(filePath, slideIndex),
            "duplicate" => powerPointService.DuplicateSlide(filePath, slideIndex),
            "reorder" when targetIndex.HasValue => powerPointService.ReorderSlide(filePath, slideIndex, targetIndex.Value),
            "reorder" => new DocumentResult(false, "Target index is required for reorder operation"),
            _ => new DocumentResult(false, $"Unknown operation: {operation}. Use 'delete', 'duplicate', or 'reorder'.")
        };
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Perform batch operations on a PowerPoint presentation using a JSON array.

**Operations JSON array** - each object has a 'type' and type-specific properties:

| Type | Key Properties |
|------|----------------|
| addSlide | backgroundColor |
| deleteSlide | slideIndex |
| duplicateSlide | sourceIndex |
| reorderSlide | fromIndex, toIndex |
| setBackground | slideIndex, backgroundColor |
| addTitle | slideIndex, title, subtitle, bold, fontSize, fontColor |
| addTextBox | slideIndex, text, xInches, yInches, widthInches, heightInches |
| addBulletPoints | slideIndex, points (array), xInches, yInches |
| addImage | slideIndex, imagePath, xInches, yInches, widthInches, heightInches |
| addShape | slideIndex, shapeType, xInches, yInches, fillColor, borderColor |
| addTable | slideIndex, tableData (2D array), xInches, yInches |
| addSpeakerNotes | slideIndex, notes |

**Shape types**: rectangle, roundRectangle, ellipse, triangle, diamond, pentagon, hexagon, arrow, leftArrow, star, heart")]
    public string BatchModifyPowerPointPresentation(
        [Description("Path to the existing PowerPoint presentation")] string filePath,
        [Description("JSON array of operations")] string operationsJson)
    {
        try
        {
            var operations = JsonSerializer.Deserialize<PowerPointOperation[]>(operationsJson, JsonOptions);
            if (operations == null || operations.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "No operations provided"), JsonOptions);
            }

            var details = new List<OperationOutcome>();
            int successCount = 0;
            int failCount = 0;

            for (int i = 0; i < operations.Length; i++)
            {
                var op = operations[i];
                DocumentResult? opResult = null;

                try
                {
                    opResult = op.Type.ToLowerInvariant() switch
                    {
                        "addslide" => ProcessAddSlide(filePath, op),
                        "deleteslide" => ProcessDeleteSlide(filePath, op),
                        "duplicateslide" => ProcessDuplicateSlide(filePath, op),
                        "reorderslide" => ProcessReorderSlide(filePath, op),
                        "setbackground" => ProcessSetBackground(filePath, op),
                        "addtitle" => ProcessAddTitle(filePath, op),
                        "addtextbox" => ProcessAddTextBox(filePath, op),
                        "addbulletpoints" => ProcessAddBulletPoints(filePath, op),
                        "addimage" => ProcessAddImage(filePath, op),
                        "addshape" => ProcessAddShape(filePath, op),
                        "addtable" => ProcessAddTable(filePath, op),
                        "addspeakernotes" => ProcessAddSpeakerNotes(filePath, op),
                        _ => new DocumentResult(false, $"Unknown operation type: {op.Type}")
                    };

                    if (opResult.Success)
                    {
                        successCount++;
                        details.Add(new OperationOutcome(i, op.Type, true, "Success"));
                    }
                    else
                    {
                        failCount++;
                        details.Add(new OperationOutcome(i, op.Type, false, opResult.Message));
                    }
                }
                catch (Exception ex)
                {
                    failCount++;
                    details.Add(new OperationOutcome(i, op.Type, false, ex.Message));
                }
            }

            var batchResult = new BatchOperationResult(
                Success: failCount == 0,
                Message: failCount == 0
                    ? $"All {successCount} operations completed successfully"
                    : $"{successCount} succeeded, {failCount} failed",
                TotalOperations: operations.Length,
                SuccessfulOperations: successCount,
                FailedOperations: failCount,
                Details: details
            );

            return JsonSerializer.Serialize(batchResult, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description(@"Read content from a PowerPoint presentation.

**Options for 'readType' parameter**:
- 'all' (default): Get all text from all slides
- 'slide': Get text from a specific slide (use slideIndex)
- 'count': Get the total number of slides

**Examples**:
- Get all text: {""filePath"": ""C:/presentations/demo.pptx""}
- Get slide 2: {""filePath"": ""C:/presentations/demo.pptx"", ""readType"": ""slide"", ""slideIndex"": 2}
- Get slide count: {""filePath"": ""C:/presentations/demo.pptx"", ""readType"": ""count""}")]
    public string ReadPowerPointPresentation(
        [Description("Path to the PowerPoint presentation")] string filePath,
        [Description("Type of read: 'all', 'slide', or 'count'")] string readType = "all",
        [Description("Slide index for 'slide' read type (0-based)")] int? slideIndex = null)
    {
        ContentResult result = readType.ToLowerInvariant() switch
        {
            "slide" when slideIndex.HasValue => powerPointService.GetSlideText(filePath, slideIndex.Value),
            "count" => powerPointService.GetSlideCount(filePath),
            _ => powerPointService.GetAllSlidesText(filePath)
        };

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    #region Private Operation Processors

    private DocumentResult ProcessAddSlide(string filePath, PowerPointOperation op)
    {
        var layoutOptions = new SlideLayoutOptions(BackgroundColor: op.BackgroundColor);
        return powerPointService.AddSlide(filePath, layoutOptions);
    }

    private DocumentResult ProcessDeleteSlide(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue)
        {
            return new DocumentResult(false, "Slide index is required");
        }
        return powerPointService.DeleteSlide(filePath, op.SlideIndex.Value);
    }

    private DocumentResult ProcessDuplicateSlide(string filePath, PowerPointOperation op)
    {
        if (!op.SourceIndex.HasValue)
        {
            return new DocumentResult(false, "Source index is required");
        }
        return powerPointService.DuplicateSlide(filePath, op.SourceIndex.Value);
    }

    private DocumentResult ProcessReorderSlide(string filePath, PowerPointOperation op)
    {
        if (!op.FromIndex.HasValue || !op.ToIndex.HasValue)
        {
            return new DocumentResult(false, "From index and to index are required");
        }
        return powerPointService.ReorderSlide(filePath, op.FromIndex.Value, op.ToIndex.Value);
    }

    private DocumentResult ProcessSetBackground(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.BackgroundColor))
        {
            return new DocumentResult(false, "Slide index and background color are required");
        }
        return powerPointService.SetSlideBackground(filePath, op.SlideIndex.Value, op.BackgroundColor);
    }

    private DocumentResult ProcessAddTitle(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.Title))
        {
            return new DocumentResult(false, "Slide index and title are required");
        }
        var textFormat = new TextFormatting(
            Bold: op.Bold ?? true,
            FontSize: op.FontSize ?? 44,
            FontColor: op.FontColor
        );
        return powerPointService.AddTitle(filePath, op.SlideIndex.Value, op.Title, op.Subtitle, textFormat);
    }

    private DocumentResult ProcessAddTextBox(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.Text))
        {
            return new DocumentResult(false, "Slide index and text are required");
        }
        var textFormat = new TextFormatting(
            Bold: op.Bold ?? false,
            FontSize: op.FontSize ?? 18,
            FontColor: op.FontColor
        );
        var options = new TextBoxOptions(
            X: (long)((op.XInches ?? 1.0) * 914400),
            Y: (long)((op.YInches ?? 2.0) * 914400),
            Width: (long)((op.WidthInches ?? 8.0) * 914400),
            Height: (long)((op.HeightInches ?? 1.0) * 914400),
            BackgroundColor: op.BackgroundColor,
            BorderColor: op.BorderColor,
            TextFormat: textFormat
        );
        return powerPointService.AddTextBox(filePath, op.SlideIndex.Value, op.Text, options);
    }

    private DocumentResult ProcessAddBulletPoints(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || op.Points == null || op.Points.Length == 0)
        {
            return new DocumentResult(false, "Slide index and points array are required");
        }
        var textFormat = new TextFormatting(
            FontSize: op.FontSize ?? 18,
            FontColor: op.FontColor
        );
        var options = new TextBoxOptions(
            X: (long)((op.XInches ?? 1.0) * 914400),
            Y: (long)((op.YInches ?? 2.0) * 914400),
            Width: (long)((op.WidthInches ?? 8.0) * 914400),
            Height: (long)((op.HeightInches ?? 4.0) * 914400),
            TextFormat: textFormat
        );
        return powerPointService.AddBulletPoints(filePath, op.SlideIndex.Value, op.Points, options);
    }

    private DocumentResult ProcessAddImage(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.ImagePath))
        {
            return new DocumentResult(false, "Slide index and image path are required");
        }
        var options = new ImageOptions(
            WidthEmu: (long)((op.WidthInches ?? 4.0) * 914400),
            HeightEmu: (long)((op.HeightInches ?? 3.0) * 914400),
            AltText: op.AltText
        );
        return powerPointService.AddImage(
            filePath,
            op.SlideIndex.Value,
            op.ImagePath,
            (long)((op.XInches ?? 1.0) * 914400),
            (long)((op.YInches ?? 2.0) * 914400),
            options
        );
    }

    private DocumentResult ProcessAddShape(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.ShapeType))
        {
            return new DocumentResult(false, "Slide index and shape type are required");
        }
        var options = new ShapeOptions(
            ShapeType: op.ShapeType,
            X: (long)((op.XInches ?? 1.0) * 914400),
            Y: (long)((op.YInches ?? 2.0) * 914400),
            Width: (long)((op.WidthInches ?? 2.0) * 914400),
            Height: (long)((op.HeightInches ?? 2.0) * 914400),
            FillColor: op.FillColor,
            BorderColor: op.BorderColor,
            BorderWidth: op.BorderWidth ?? 1.0
        );
        return powerPointService.AddShape(filePath, op.SlideIndex.Value, options);
    }

    private DocumentResult ProcessAddTable(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || op.TableData == null || op.TableData.Length == 0)
        {
            return new DocumentResult(false, "Slide index and table data are required");
        }
        return powerPointService.AddTable(
            filePath,
            op.SlideIndex.Value,
            op.TableData,
            (long)((op.XInches ?? 1.0) * 914400),
            (long)((op.YInches ?? 2.0) * 914400),
            (long)((op.WidthInches ?? 8.0) * 914400),
            (long)((op.HeightInches ?? 3.0) * 914400)
        );
    }

    private DocumentResult ProcessAddSpeakerNotes(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.Notes))
        {
            return new DocumentResult(false, "Slide index and notes are required");
        }
        return powerPointService.AddSpeakerNotes(filePath, op.SlideIndex.Value, op.Notes);
    }

    #endregion
}
