using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// AI-Optimized MCP Tools for PowerPoint presentations.
/// Consolidated tools with simplified descriptions and tool annotations for better AI discoverability.
/// </summary>
[McpServerToolType]
public sealed class PowerPointDocumentToolsOptimized(IPowerPointDocumentService powerPointService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    #region Core Presentation Operations

    [McpServerTool(Name = "pptx_create", Destructive = false, ReadOnly = false), Description("Creates a PowerPoint presentation (.pptx) with optional title slide.")]
    public string CreatePowerPointPresentation(
        [Description("Full path (e.g., C:/slides/demo.pptx)")] string filePath,
        [Description("Title for first slide")] string? title = null,
        [Description("Subtitle for title slide")] string? subtitle = null)
    {
        var result = powerPointService.CreatePresentation(filePath, title);
        if (!result.Success)
            return JsonSerializer.Serialize(result, JsonOptions);

        if (!string.IsNullOrWhiteSpace(subtitle) && !string.IsNullOrWhiteSpace(title))
        {
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

    [McpServerTool(Name = "pptx_read", Destructive = false, ReadOnly = true), Description("Reads text from a PowerPoint presentation. Returns all slides by default.")]
    public string ReadPowerPointPresentation(
        [Description("Path to the presentation")] string filePath,
        [Description("all (default), slide, or count")] string readType = "all",
        [Description("Slide index (0-based) for 'slide' mode")] int? slideIndex = null)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new ContentResult(false, null, $"File not found: {filePath}. Use pptx_create to create a presentation first."), JsonOptions);

        ContentResult result = readType.ToLowerInvariant() switch
        {
            "slide" when slideIndex.HasValue => powerPointService.GetSlideText(filePath, slideIndex.Value),
            "slide" => new ContentResult(false, null, "slideIndex is required when readType is 'slide'"),
            "count" => powerPointService.GetSlideCount(filePath),
            _ => powerPointService.GetAllSlidesText(filePath)
        };

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    #endregion

    #region Slide Operations

    [McpServerTool(Name = "pptx_add_slide", Destructive = false, ReadOnly = false), Description("Adds a new blank slide to a presentation.")]
    public string AddPowerPointSlide(
        [Description("Path to the presentation")] string filePath,
        [Description("Background color as hex (e.g., FFFFFF)")] string? backgroundColor = null)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {filePath}", Suggestion: "Use pptx_create to create the presentation first"), JsonOptions);

        var layoutOptions = new SlideLayoutOptions(BackgroundColor: backgroundColor);
        var result = powerPointService.AddSlide(filePath, layoutOptions);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "pptx_manage_slide", Destructive = true, ReadOnly = false), Description("Manages slides: delete, duplicate, or reorder.")]
    public string ManagePowerPointSlide(
        [Description("Path to the presentation")] string filePath,
        [Description("delete, duplicate, or reorder")] string operation,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Target index for reorder")] int? targetIndex = null)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {filePath}", Suggestion: "Use pptx_create to create the presentation first"), JsonOptions);

        DocumentResult result = operation.ToLowerInvariant() switch
        {
            "delete" => powerPointService.DeleteSlide(filePath, slideIndex),
            "duplicate" => powerPointService.DuplicateSlide(filePath, slideIndex),
            "reorder" when targetIndex.HasValue => powerPointService.ReorderSlide(filePath, slideIndex, targetIndex.Value),
            "reorder" => new DocumentResult(false, "Target index required for reorder", Suggestion: "Provide the 'targetIndex' parameter"),
            _ => new DocumentResult(false, $"Unknown operation: '{operation}'", Suggestion: "Valid operations: delete, duplicate, reorder")
        };
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    #endregion

    #region Content Operations

    [McpServerTool(Name = "pptx_add_title", Destructive = false, ReadOnly = false), Description("Adds a title and optional subtitle to a slide.")]
    public string AddPowerPointTitle(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Title text")] string title,
        [Description("Subtitle text")] string? subtitle = null,
        [Description("Bold text")] bool bold = true,
        [Description("Font size in points")] int fontSize = 44,
        [Description("Font color as hex")] string? fontColor = null)
    {
        var textFormat = new TextFormatting(Bold: bold, FontSize: fontSize, FontColor: fontColor);
        var result = powerPointService.AddTitle(filePath, slideIndex, title, subtitle, textFormat);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "pptx_add_text", Destructive = false, ReadOnly = false), Description("Adds text or bullet points to a slide.")]
    public string AddPowerPointText(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Text content or first bullet")] string text,
        [Description("Additional bullets as JSON array: [\"Point 2\", \"Point 3\"]")] string? additionalPointsJson = null,
        [Description("X position in inches")] double xInches = 1.0,
        [Description("Y position in inches")] double yInches = 2.0,
        [Description("Width in inches")] double widthInches = 8.0,
        [Description("Height in inches")] double heightInches = 1.0,
        [Description("Font size")] int fontSize = 18,
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

        if (!string.IsNullOrWhiteSpace(additionalPointsJson))
        {
            try
            {
                var additionalPoints = JsonSerializer.Deserialize<string[]>(additionalPointsJson, JsonOptions) ?? [];
                var allPoints = new[] { text }.Concat(additionalPoints).ToArray();
                var result = powerPointService.AddBulletPoints(filePath, slideIndex, allPoints, options);
                return JsonSerializer.Serialize(result, JsonOptions);
            }
            catch (JsonException ex)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON: {ex.Message}"), JsonOptions);
            }
        }

        return JsonSerializer.Serialize(powerPointService.AddTextBox(filePath, slideIndex, text, options), JsonOptions);
    }

    [McpServerTool(Name = "pptx_add_image", Destructive = false, ReadOnly = false), Description("Adds an image to a slide.")]
    public string AddPowerPointImage(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Path to image file")] string imagePath,
        [Description("X position in inches")] double xInches = 1.0,
        [Description("Y position in inches")] double yInches = 2.0,
        [Description("Width in inches")] double widthInches = 4.0,
        [Description("Height in inches")] double heightInches = 3.0,
        [Description("Alt text for accessibility")] string? altText = null)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"Presentation not found: {filePath}", Suggestion: "Use pptx_create to create the presentation first"), JsonOptions);

        if (!File.Exists(imagePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"Image not found: {imagePath}", Suggestion: "Verify the image path exists and is accessible"), JsonOptions);

        var options = new ImageOptions(
            WidthEmu: (long)(widthInches * 914400),
            HeightEmu: (long)(heightInches * 914400),
            AltText: altText
        );
        var result = powerPointService.AddImage(filePath, slideIndex, imagePath, (long)(xInches * 914400), (long)(yInches * 914400), options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "pptx_add_table", Destructive = false, ReadOnly = false), Description("Adds a table to a slide.")]
    public string AddPowerPointTable(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Table data as JSON 2D array: [[\"Header1\",\"Header2\"],[\"Val1\",\"Val2\"]]")] string tableDataJson,
        [Description("X position in inches")] double xInches = 1.0,
        [Description("Y position in inches")] double yInches = 2.0,
        [Description("Width in inches")] double widthInches = 8.0,
        [Description("Height in inches")] double heightInches = 3.0)
    {
        try
        {
            var tableData = JsonSerializer.Deserialize<string[][]>(tableDataJson, JsonOptions);
            if (tableData == null || tableData.Length == 0)
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty table data"), JsonOptions);

            var result = powerPointService.AddTable(
                filePath, slideIndex, tableData,
                (long)(xInches * 914400), (long)(yInches * 914400),
                (long)(widthInches * 914400), (long)(heightInches * 914400)
            );
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool(Name = "pptx_add_notes", Destructive = false, ReadOnly = false), Description("Adds speaker notes to a slide.")]
    public string AddPowerPointSpeakerNotes(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Speaker notes text")] string notes)
    {
        var result = powerPointService.AddSpeakerNotes(filePath, slideIndex, notes);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    #endregion

    #region Batch Operations

    [McpServerTool(Name = "pptx_batch", Destructive = true, ReadOnly = false), Description("Performs multiple operations. Types: addSlide, deleteSlide, duplicateSlide, reorderSlide, setBackground, addTitle, addTextBox, addBulletPoints, addImage, addShape, addTable, addSpeakerNotes.")]
    public string BatchModifyPowerPointPresentation(
        [Description("Path to the presentation")] string filePath,
        [Description("JSON array: [{\"type\":\"addSlide\"}, {\"type\":\"addTitle\",\"slideIndex\":1,\"title\":\"Intro\"}]")] string operationsJson)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {filePath}", Suggestion: "Use pptx_create to create the presentation first"), JsonOptions);

        try
        {
            var operations = JsonSerializer.Deserialize<PowerPointOperation[]>(operationsJson, JsonOptions);
            if (operations == null || operations.Length == 0)
                return JsonSerializer.Serialize(new DocumentResult(false, "No operations provided", Suggestion: "Provide a JSON array of operations"), JsonOptions);

            var details = new List<OperationOutcome>();
            int successCount = 0, failCount = 0;

            for (int i = 0; i < operations.Length; i++)
            {
                var op = operations[i];
                try
                {
                    var opResult = op.Type.ToLowerInvariant() switch
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
                        _ => new DocumentResult(false, $"Unknown type: '{op.Type}'", Suggestion: "Valid types: addSlide, deleteSlide, duplicateSlide, reorderSlide, setBackground, addTitle, addTextBox, addBulletPoints, addImage, addShape, addTable, addSpeakerNotes")
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

            return JsonSerializer.Serialize(new BatchOperationResult(
                Success: failCount == 0,
                Message: failCount == 0 ? $"All {successCount} operations completed" : $"{successCount} succeeded, {failCount} failed",
                TotalOperations: operations.Length,
                SuccessfulOperations: successCount,
                FailedOperations: failCount,
                Details: details
            ), JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON: {ex.Message}", Suggestion: "Ensure operationsJson is a valid JSON array"), JsonOptions);
        }
    }

    #endregion

    #region Private Operation Processors

    private DocumentResult ProcessAddSlide(string filePath, PowerPointOperation op)
    {
        var layoutOptions = new SlideLayoutOptions(BackgroundColor: op.BackgroundColor);
        return powerPointService.AddSlide(filePath, layoutOptions);
    }

    private DocumentResult ProcessDeleteSlide(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue)
            return new DocumentResult(false, "Slide index is required");
        return powerPointService.DeleteSlide(filePath, op.SlideIndex.Value);
    }

    private DocumentResult ProcessDuplicateSlide(string filePath, PowerPointOperation op)
    {
        if (!op.SourceIndex.HasValue)
            return new DocumentResult(false, "Source index is required");
        return powerPointService.DuplicateSlide(filePath, op.SourceIndex.Value);
    }

    private DocumentResult ProcessReorderSlide(string filePath, PowerPointOperation op)
    {
        if (!op.FromIndex.HasValue || !op.ToIndex.HasValue)
            return new DocumentResult(false, "From and to indices are required");
        return powerPointService.ReorderSlide(filePath, op.FromIndex.Value, op.ToIndex.Value);
    }

    private DocumentResult ProcessSetBackground(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.BackgroundColor))
            return new DocumentResult(false, "Slide index and background color are required");
        return powerPointService.SetSlideBackground(filePath, op.SlideIndex.Value, op.BackgroundColor);
    }

    private DocumentResult ProcessAddTitle(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.Title))
            return new DocumentResult(false, "Slide index and title are required");
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
            return new DocumentResult(false, "Slide index and text are required");
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
            return new DocumentResult(false, "Slide index and points are required");
        var textFormat = new TextFormatting(FontSize: op.FontSize ?? 18, FontColor: op.FontColor);
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
            return new DocumentResult(false, "Slide index and image path are required");
        var options = new ImageOptions(
            WidthEmu: (long)((op.WidthInches ?? 4.0) * 914400),
            HeightEmu: (long)((op.HeightInches ?? 3.0) * 914400),
            AltText: op.AltText
        );
        return powerPointService.AddImage(filePath, op.SlideIndex.Value, op.ImagePath,
            (long)((op.XInches ?? 1.0) * 914400), (long)((op.YInches ?? 2.0) * 914400), options);
    }

    private DocumentResult ProcessAddShape(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.ShapeType))
            return new DocumentResult(false, "Slide index and shape type are required");
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
            return new DocumentResult(false, "Slide index and table data are required");
        return powerPointService.AddTable(filePath, op.SlideIndex.Value, op.TableData,
            (long)((op.XInches ?? 1.0) * 914400), (long)((op.YInches ?? 2.0) * 914400),
            (long)((op.WidthInches ?? 8.0) * 914400), (long)((op.HeightInches ?? 3.0) * 914400));
    }

    private DocumentResult ProcessAddSpeakerNotes(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.Notes))
            return new DocumentResult(false, "Slide index and notes are required");
        return powerPointService.AddSpeakerNotes(filePath, op.SlideIndex.Value, op.Notes);
    }

    #endregion
}
