using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// LEGACY: Format-specific PowerPoint tools - kept for backward compatibility.
/// Use the unified office_* tools from OfficeDocumentToolsConsolidated instead.
/// This class is no longer registered as an MCP tool provider.
/// </summary>
// [McpServerToolType] - Disabled: Use consolidated office_* tools instead
public sealed class PowerPointDocumentToolsOptimized(IPowerPointDocumentService powerPointService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    #region Core Presentation Operations

    // [McpServerTool] - Disabled: Use office_create instead
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

    [McpServerTool(Name = "pptx_add_text", Destructive = false, ReadOnly = false), Description("Adds a text box to a slide. Supports single text or rich multi-run text with different formatting per run.")]
    public string AddPowerPointText(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Text content (for simple text box)")] string? text = null,
        [Description("Rich text paragraphs as JSON: [{\"runs\":[{\"text\":\"Bold \",\"format\":{\"bold\":true}},{\"text\":\"normal\"}],\"alignment\":\"Center\"}]")] string? paragraphsJson = null,
        [Description("Additional bullets as JSON array: [\"Point 2\", \"Point 3\"]")] string? additionalPointsJson = null,
        [Description("X position in inches")] double xInches = 1.0,
        [Description("Y position in inches")] double yInches = 2.0,
        [Description("Width in inches")] double widthInches = 8.0,
        [Description("Height in inches")] double heightInches = 1.0,
        [Description("Font size")] int fontSize = 18,
        [Description("Font color as hex")] string? fontColor = null,
        [Description("Font name (e.g., Segoe UI, Calibri)")] string? fontName = null,
        [Description("Bold text")] bool bold = false,
        [Description("Italic text")] bool italic = false,
        [Description("Horizontal alignment: Left, Center, Right, Justify")] string alignment = "Left",
        [Description("Vertical alignment: Top, Middle, Bottom")] string verticalAlignment = "Top",
        [Description("Background color as hex")] string? backgroundColor = null,
        [Description("Border color as hex")] string? borderColor = null,
        [Description("Border width in points")] double borderWidth = 1.0,
        [Description("Word wrap text")] bool wordWrap = true,
        [Description("Auto fit: None, ShrinkText, ResizeShape")] string autoFit = "None",
        [Description("Rotation in degrees")] double rotation = 0.0,
        [Description("Gradient fill as JSON: {\"stops\":[{\"position\":0,\"color\":\"FF0000\"},{\"position\":100,\"color\":\"0000FF\"}],\"angle\":90}")] string? gradientJson = null)
    {
        var textFormat = new TextFormatting(Bold: bold, Italic: italic, FontSize: fontSize, FontColor: fontColor, FontName: fontName);
        
        GradientFillOptions? gradient = null;
        if (!string.IsNullOrWhiteSpace(gradientJson))
        {
            try { gradient = JsonSerializer.Deserialize<GradientFillOptions>(gradientJson, JsonOptions); }
            catch (JsonException ex) { return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid gradient JSON: {ex.Message}"), JsonOptions); }
        }

        RichParagraph[]? paragraphs = null;
        if (!string.IsNullOrWhiteSpace(paragraphsJson))
        {
            try { paragraphs = JsonSerializer.Deserialize<RichParagraph[]>(paragraphsJson, JsonOptions); }
            catch (JsonException ex) { return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid paragraphs JSON: {ex.Message}"), JsonOptions); }
        }

        var options = new TextBoxOptions(
            X: (long)(xInches * 914400),
            Y: (long)(yInches * 914400),
            Width: (long)(widthInches * 914400),
            Height: (long)(heightInches * 914400),
            TextFormat: textFormat,
            Alignment: alignment,
            VerticalAlignment: verticalAlignment,
            BackgroundColor: backgroundColor,
            BorderColor: borderColor,
            BorderWidth: borderWidth,
            WordWrap: wordWrap,
            AutoFit: autoFit,
            Rotation: rotation,
            Paragraphs: paragraphs,
            GradientFill: gradient
        );

        // Rich text mode
        if (paragraphs != null && paragraphs.Length > 0)
        {
            var result = powerPointService.AddRichTextBox(filePath, slideIndex, options);
            return JsonSerializer.Serialize(result, JsonOptions);
        }

        // Bullet points mode
        if (!string.IsNullOrWhiteSpace(additionalPointsJson))
        {
            try
            {
                var additionalPoints = JsonSerializer.Deserialize<string[]>(additionalPointsJson, JsonOptions) ?? [];
                var allPoints = new[] { text ?? string.Empty }.Concat(additionalPoints).ToArray();
                var result = powerPointService.AddBulletPoints(filePath, slideIndex, allPoints, options);
                return JsonSerializer.Serialize(result, JsonOptions);
            }
            catch (JsonException ex)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON: {ex.Message}"), JsonOptions);
            }
        }

        // Simple text mode
        return JsonSerializer.Serialize(powerPointService.AddTextBox(filePath, slideIndex, text ?? string.Empty, options), JsonOptions);
    }

    [McpServerTool(Name = "pptx_add_image", Destructive = false, ReadOnly = false), Description("Adds an image to a slide. Supports JPEG, PNG, GIF, BMP.")]
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

    [McpServerTool(Name = "pptx_add_shape", Destructive = false, ReadOnly = false), Description("Adds a shape to a slide with optional text, gradient fill, shadow, rotation, transparency, and custom corner radius. Supports 100+ shape types including rectangle, roundRectangle, ellipse, triangle, diamond, pentagon, hexagon, chevron, arrow types, stars, hearts, clouds, flowchart shapes, callouts, braces, brackets, math symbols, gears, and more.")]
    public string AddPowerPointShape(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Shape type: rectangle, roundRectangle, ellipse, triangle, diamond, pentagon, hexagon, chevron, rightArrow, leftArrow, star5, heart, cloud, leftBrace, rightBrace, leftBracket, rightBracket, flowChartProcess, flowChartDecision, callout1, donut, cube, ribbon, plus, and many more")] string shapeType,
        [Description("X position in inches")] double xInches = 1.0,
        [Description("Y position in inches")] double yInches = 2.0,
        [Description("Width in inches")] double widthInches = 2.0,
        [Description("Height in inches")] double heightInches = 2.0,
        [Description("Fill color as hex (e.g., 4472C4)")] string? fillColor = null,
        [Description("Border color as hex")] string? borderColor = null,
        [Description("Border width in points")] double borderWidth = 1.0,
        [Description("Text to display inside the shape")] string? text = null,
        [Description("Font size for shape text")] int fontSize = 14,
        [Description("Font color as hex for shape text")] string? fontColor = null,
        [Description("Font name for shape text")] string? fontName = null,
        [Description("Bold shape text")] bool bold = false,
        [Description("Italic shape text")] bool italic = false,
        [Description("Text horizontal alignment: Left, Center, Right")] string textAlignment = "Center",
        [Description("Text vertical alignment: Top, Middle, Bottom")] string verticalTextAlignment = "Middle",
        [Description("Rotation in degrees")] double rotation = 0.0,
        [Description("Corner radius in points (for roundRectangle)")] int? cornerRadiusPt = null,
        [Description("Add drop shadow")] bool hasShadow = false,
        [Description("Border dash style: Solid, Dash, Dot, DashDot, LongDash")] string dashStyle = "Solid",
        [Description("Fill transparency 0-100")] int transparencyPercent = 0,
        [Description("Gradient fill JSON: {\"stops\":[{\"position\":0,\"color\":\"FF0000\"},{\"position\":100,\"color\":\"0000FF\"}],\"angle\":90}")] string? gradientJson = null,
        [Description("Rich text paragraphs JSON (overrides text param)")] string? paragraphsJson = null,
        [Description("No fill (transparent shape)")] bool noFill = false)
    {
        GradientFillOptions? gradient = null;
        if (!string.IsNullOrWhiteSpace(gradientJson))
        {
            try { gradient = JsonSerializer.Deserialize<GradientFillOptions>(gradientJson, JsonOptions); }
            catch (JsonException ex) { return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid gradient JSON: {ex.Message}"), JsonOptions); }
        }

        RichParagraph[]? paragraphs = null;
        if (!string.IsNullOrWhiteSpace(paragraphsJson))
        {
            try { paragraphs = JsonSerializer.Deserialize<RichParagraph[]>(paragraphsJson, JsonOptions); }
            catch (JsonException ex) { return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid paragraphs JSON: {ex.Message}"), JsonOptions); }
        }

        var textFormat = new TextFormatting(Bold: bold, Italic: italic, FontSize: fontSize, FontColor: fontColor, FontName: fontName);
        var options = new ShapeOptions(
            ShapeType: shapeType,
            X: (long)(xInches * 914400),
            Y: (long)(yInches * 914400),
            Width: (long)(widthInches * 914400),
            Height: (long)(heightInches * 914400),
            FillColor: fillColor,
            BorderColor: borderColor,
            BorderWidth: borderWidth,
            Text: text,
            TextFormat: textFormat,
            TextAlignment: textAlignment,
            VerticalTextAlignment: verticalTextAlignment,
            Rotation: rotation,
            CornerRadiusPt: cornerRadiusPt,
            HasShadow: hasShadow,
            DashStyle: dashStyle,
            TransparencyPercent: transparencyPercent,
            GradientFill: gradient,
            Paragraphs: paragraphs,
            NoFill: noFill
        );
        var result = powerPointService.AddShape(filePath, slideIndex, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "pptx_add_line", Destructive = false, ReadOnly = false), Description("Adds a line to a slide. Supports arrow heads, dash styles, and custom thickness.")]
    public string AddPowerPointLine(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Start X position in inches")] double x1Inches,
        [Description("Start Y position in inches")] double y1Inches,
        [Description("End X position in inches")] double x2Inches,
        [Description("End Y position in inches")] double y2Inches,
        [Description("Line color as hex (default: black)")] string? lineColor = null,
        [Description("Line width in points")] double lineWidth = 1.0,
        [Description("Dash style: Solid, Dash, Dot, DashDot, LongDash")] string dashStyle = "Solid",
        [Description("Start arrow type: None, Triangle, Stealth, Diamond, Oval, Open")] string? startArrow = null,
        [Description("End arrow type: None, Triangle, Stealth, Diamond, Oval, Open")] string? endArrow = null)
    {
        var options = new LineOptions(
            X1: (long)(x1Inches * 914400),
            Y1: (long)(y1Inches * 914400),
            X2: (long)(x2Inches * 914400),
            Y2: (long)(y2Inches * 914400),
            LineColor: lineColor,
            LineWidth: lineWidth,
            DashStyle: dashStyle,
            StartArrow: startArrow,
            EndArrow: endArrow
        );
        var result = powerPointService.AddLine(filePath, slideIndex, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "pptx_add_connector", Destructive = false, ReadOnly = false), Description("Adds a connector (straight, elbow, or curved) between two points on a slide. Supports arrow heads and dash styles.")]
    public string AddPowerPointConnector(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Start X position in inches")] double x1Inches,
        [Description("Start Y position in inches")] double y1Inches,
        [Description("End X position in inches")] double x2Inches,
        [Description("End Y position in inches")] double y2Inches,
        [Description("Connector type: Straight, Elbow, Curved")] string connectorType = "Straight",
        [Description("Line color as hex")] string? lineColor = null,
        [Description("Line width in points")] double lineWidth = 1.0,
        [Description("Dash style: Solid, Dash, Dot, DashDot, LongDash")] string dashStyle = "Solid",
        [Description("Start arrow type: None, Triangle, Stealth, Diamond, Oval, Open")] string? startArrow = null,
        [Description("End arrow type: None, Triangle, Stealth, Diamond, Oval, Open")] string? endArrow = null)
    {
        var options = new ConnectorOptions(
            X1: (long)(x1Inches * 914400),
            Y1: (long)(y1Inches * 914400),
            X2: (long)(x2Inches * 914400),
            Y2: (long)(y2Inches * 914400),
            ConnectorType: connectorType,
            LineColor: lineColor,
            LineWidth: lineWidth,
            DashStyle: dashStyle,
            StartArrow: startArrow,
            EndArrow: endArrow
        );
        var result = powerPointService.AddConnector(filePath, slideIndex, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "pptx_add_group", Destructive = false, ReadOnly = false), Description("Adds a group of shapes to a slide. Items inside the group are positioned relative to the group's coordinate space.")]
    public string AddPowerPointGroupShape(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Group X position in inches")] double xInches,
        [Description("Group Y position in inches")] double yInches,
        [Description("Group width in inches")] double widthInches,
        [Description("Group height in inches")] double heightInches,
        [Description("Group items as JSON: [{\"itemType\":\"shape\",\"x\":0,\"y\":0,\"width\":914400,\"height\":914400,\"shapeType\":\"rectangle\",\"fillColor\":\"FF0000\",\"text\":\"Box 1\"},{\"itemType\":\"textbox\",\"x\":914400,\"y\":0,\"width\":914400,\"height\":914400,\"text\":\"Label\"}]")] string groupItemsJson)
    {
        try
        {
            var items = JsonSerializer.Deserialize<GroupShapeItem[]>(groupItemsJson, JsonOptions);
            if (items == null || items.Length == 0)
                return JsonSerializer.Serialize(new DocumentResult(false, "No group items provided"), JsonOptions);

            var result = powerPointService.AddGroupShape(
                filePath, slideIndex,
                (long)(xInches * 914400), (long)(yInches * 914400),
                (long)(widthInches * 914400), (long)(heightInches * 914400),
                items);
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool(Name = "pptx_set_slide_size", Destructive = false, ReadOnly = false), Description("Sets the slide size for the presentation. Affects all slides.")]
    public string SetPowerPointSlideSize(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide size: Widescreen (16:9), Standard (4:3), Widescreen16x10 (16:10), A4, Letter")] string size = "Widescreen")
    {
        var result = powerPointService.SetSlideSize(filePath, size);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "pptx_set_background", Destructive = false, ReadOnly = false), Description("Sets the background of a slide to a solid color or gradient.")]
    public string SetPowerPointBackground(
        [Description("Path to the presentation")] string filePath,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Solid background color as hex (e.g., FFFFFF)")] string? color = null,
        [Description("Gradient fill JSON: {\"stops\":[{\"position\":0,\"color\":\"003366\"},{\"position\":100,\"color\":\"99CCFF\"}],\"angle\":270}")] string? gradientJson = null)
    {
        if (!string.IsNullOrWhiteSpace(gradientJson))
        {
            try
            {
                var gradient = JsonSerializer.Deserialize<GradientFillOptions>(gradientJson, JsonOptions);
                if (gradient != null)
                {
                    var result = powerPointService.SetSlideBackgroundGradient(filePath, slideIndex, gradient);
                    return JsonSerializer.Serialize(result, JsonOptions);
                }
            }
            catch (JsonException ex)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid gradient JSON: {ex.Message}"), JsonOptions);
            }
        }

        if (!string.IsNullOrWhiteSpace(color))
        {
            var result = powerPointService.SetSlideBackground(filePath, slideIndex, color);
            return JsonSerializer.Serialize(result, JsonOptions);
        }

        return JsonSerializer.Serialize(new DocumentResult(false, "Provide either color or gradientJson"), JsonOptions);
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

    [McpServerTool(Name = "pptx_batch", Destructive = true, ReadOnly = false), Description("Performs multiple operations on a presentation in a single call. Types: addSlide, deleteSlide, duplicateSlide, reorderSlide, setBackground, setBackgroundGradient, setSlideSize, addTitle, addTextBox, addRichTextBox, addBulletPoints, addImage, addShape, addLine, addConnector, addGroupShape, addTable, addSpeakerNotes.")]
    public string BatchModifyPowerPointPresentation(
        [Description("Path to the presentation")] string filePath,
        [Description("JSON array of operations, e.g.: [{\"type\":\"addSlide\"}, {\"type\":\"addShape\",\"slideIndex\":0,\"shapeType\":\"roundRectangle\",\"xInches\":1,\"yInches\":1,\"widthInches\":3,\"heightInches\":1,\"fillColor\":\"4472C4\",\"text\":\"Hello\",\"fontColor\":\"FFFFFF\",\"bold\":true}]")] string operationsJson)
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
                        "setbackgroundgradient" => ProcessSetBackgroundGradient(filePath, op),
                        "setslidesize" => ProcessSetSlideSize(filePath, op),
                        "addtitle" => ProcessAddTitle(filePath, op),
                        "addtextbox" => ProcessAddTextBox(filePath, op),
                        "addrichtextbox" => ProcessAddRichTextBox(filePath, op),
                        "addbulletpoints" => ProcessAddBulletPoints(filePath, op),
                        "addimage" => ProcessAddImage(filePath, op),
                        "addshape" => ProcessAddShape(filePath, op),
                        "addline" => ProcessAddLine(filePath, op),
                        "addconnector" => ProcessAddConnector(filePath, op),
                        "addgroupshape" => ProcessAddGroupShape(filePath, op),
                        "addtable" => ProcessAddTable(filePath, op),
                        "addspeakernotes" => ProcessAddSpeakerNotes(filePath, op),
                        _ => new DocumentResult(false, $"Unknown type: '{op.Type}'", Suggestion: "Valid types: addSlide, deleteSlide, duplicateSlide, reorderSlide, setBackground, setBackgroundGradient, setSlideSize, addTitle, addTextBox, addRichTextBox, addBulletPoints, addImage, addShape, addLine, addConnector, addGroupShape, addTable, addSpeakerNotes")
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
            Italic: op.Italic ?? false,
            Underline: op.Underline ?? false,
            FontSize: op.FontSize ?? 18,
            FontColor: op.FontColor,
            FontName: op.FontName
        );

        GradientFillOptions? gradient = ParseGradient(op.GradientJson);

        var options = new TextBoxOptions(
            X: (long)((op.XInches ?? 1.0) * 914400),
            Y: (long)((op.YInches ?? 2.0) * 914400),
            Width: (long)((op.WidthInches ?? 8.0) * 914400),
            Height: (long)((op.HeightInches ?? 1.0) * 914400),
            BackgroundColor: op.BackgroundColor,
            BorderColor: op.BorderColor,
            BorderWidth: op.BorderWidth ?? 1.0,
            TextFormat: textFormat,
            Alignment: op.Alignment ?? "Left",
            VerticalAlignment: op.VerticalAlignment ?? "Top",
            WordWrap: op.WordWrap ?? true,
            AutoFit: op.AutoFit ?? "None",
            Rotation: op.Rotation ?? 0.0,
            GradientFill: gradient,
            MarginLeftInches: op.MarginLeftInches ?? 0.1,
            MarginRightInches: op.MarginRightInches ?? 0.1,
            MarginTopInches: op.MarginTopInches ?? 0.05,
            MarginBottomInches: op.MarginBottomInches ?? 0.05
        );
        return powerPointService.AddTextBox(filePath, op.SlideIndex.Value, op.Text, options);
    }

    private DocumentResult ProcessAddRichTextBox(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.ParagraphsJson))
            return new DocumentResult(false, "Slide index and paragraphsJson are required");

        RichParagraph[]? paragraphs;
        try { paragraphs = JsonSerializer.Deserialize<RichParagraph[]>(op.ParagraphsJson, JsonOptions); }
        catch (JsonException ex) { return new DocumentResult(false, $"Invalid paragraphs JSON: {ex.Message}"); }
        
        if (paragraphs == null || paragraphs.Length == 0)
            return new DocumentResult(false, "No paragraphs provided");

        GradientFillOptions? gradient = ParseGradient(op.GradientJson);

        var options = new TextBoxOptions(
            X: (long)((op.XInches ?? 1.0) * 914400),
            Y: (long)((op.YInches ?? 2.0) * 914400),
            Width: (long)((op.WidthInches ?? 8.0) * 914400),
            Height: (long)((op.HeightInches ?? 1.0) * 914400),
            BackgroundColor: op.BackgroundColor,
            BorderColor: op.BorderColor,
            BorderWidth: op.BorderWidth ?? 1.0,
            Alignment: op.Alignment ?? "Left",
            VerticalAlignment: op.VerticalAlignment ?? "Top",
            WordWrap: op.WordWrap ?? true,
            AutoFit: op.AutoFit ?? "None",
            Rotation: op.Rotation ?? 0.0,
            Paragraphs: paragraphs,
            GradientFill: gradient,
            MarginLeftInches: op.MarginLeftInches ?? 0.1,
            MarginRightInches: op.MarginRightInches ?? 0.1,
            MarginTopInches: op.MarginTopInches ?? 0.05,
            MarginBottomInches: op.MarginBottomInches ?? 0.05
        );
        return powerPointService.AddRichTextBox(filePath, op.SlideIndex.Value, options);
    }

    private DocumentResult ProcessAddBulletPoints(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || op.Points == null || op.Points.Length == 0)
            return new DocumentResult(false, "Slide index and points are required");
        var textFormat = new TextFormatting(FontSize: op.FontSize ?? 18, FontColor: op.FontColor, FontName: op.FontName);
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

        var textFormat = new TextFormatting(
            Bold: op.Bold ?? false,
            Italic: op.Italic ?? false,
            FontSize: op.FontSize ?? 14,
            FontColor: op.FontColor,
            FontName: op.FontName
        );

        GradientFillOptions? gradient = ParseGradient(op.GradientJson);
        RichParagraph[]? paragraphs = ParseParagraphs(op.ParagraphsJson);

        var options = new ShapeOptions(
            ShapeType: op.ShapeType,
            X: (long)((op.XInches ?? 1.0) * 914400),
            Y: (long)((op.YInches ?? 2.0) * 914400),
            Width: (long)((op.WidthInches ?? 2.0) * 914400),
            Height: (long)((op.HeightInches ?? 2.0) * 914400),
            FillColor: op.FillColor,
            BorderColor: op.BorderColor,
            BorderWidth: op.BorderWidth ?? 1.0,
            Text: op.Text,
            TextFormat: textFormat,
            TextAlignment: op.Alignment ?? "Center",
            VerticalTextAlignment: op.VerticalAlignment ?? "Middle",
            Rotation: op.Rotation ?? 0.0,
            CornerRadiusPt: op.CornerRadiusPt,
            HasShadow: op.HasShadow ?? false,
            DashStyle: op.DashStyle ?? "Solid",
            TransparencyPercent: op.TransparencyPercent ?? 0,
            GradientFill: gradient,
            Paragraphs: paragraphs,
            NoFill: op.NoFill ?? false,
            MarginLeftInches: op.MarginLeftInches ?? 0.1,
            MarginRightInches: op.MarginRightInches ?? 0.1,
            MarginTopInches: op.MarginTopInches ?? 0.05,
            MarginBottomInches: op.MarginBottomInches ?? 0.05
        );
        return powerPointService.AddShape(filePath, op.SlideIndex.Value, options);
    }

    private DocumentResult ProcessAddLine(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue)
            return new DocumentResult(false, "Slide index is required");
        var options = new LineOptions(
            X1: (long)((op.XInches ?? 0) * 914400),
            Y1: (long)((op.YInches ?? 0) * 914400),
            X2: (long)((op.X2Inches ?? 1) * 914400),
            Y2: (long)((op.Y2Inches ?? 0) * 914400),
            LineColor: op.LineColor ?? op.BorderColor,
            LineWidth: op.LineWidth ?? op.BorderWidth ?? 1.0,
            DashStyle: op.DashStyle ?? "Solid",
            StartArrow: op.StartArrow,
            EndArrow: op.EndArrow
        );
        return powerPointService.AddLine(filePath, op.SlideIndex.Value, options);
    }

    private DocumentResult ProcessAddConnector(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue)
            return new DocumentResult(false, "Slide index is required");
        var options = new ConnectorOptions(
            X1: (long)((op.XInches ?? 0) * 914400),
            Y1: (long)((op.YInches ?? 0) * 914400),
            X2: (long)((op.X2Inches ?? 1) * 914400),
            Y2: (long)((op.Y2Inches ?? 0) * 914400),
            ConnectorType: op.ConnectorType ?? "Straight",
            LineColor: op.LineColor ?? op.BorderColor,
            LineWidth: op.LineWidth ?? op.BorderWidth ?? 1.0,
            DashStyle: op.DashStyle ?? "Solid",
            StartArrow: op.StartArrow,
            EndArrow: op.EndArrow
        );
        return powerPointService.AddConnector(filePath, op.SlideIndex.Value, options);
    }

    private DocumentResult ProcessAddGroupShape(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.GroupItemsJson))
            return new DocumentResult(false, "Slide index and groupItemsJson are required");

        GroupShapeItem[]? items;
        try { items = JsonSerializer.Deserialize<GroupShapeItem[]>(op.GroupItemsJson, JsonOptions); }
        catch (JsonException ex) { return new DocumentResult(false, $"Invalid group items JSON: {ex.Message}"); }

        if (items == null || items.Length == 0)
            return new DocumentResult(false, "No group items provided");

        return powerPointService.AddGroupShape(
            filePath, op.SlideIndex.Value,
            (long)((op.XInches ?? 0) * 914400), (long)((op.YInches ?? 0) * 914400),
            (long)((op.WidthInches ?? 5) * 914400), (long)((op.HeightInches ?? 5) * 914400),
            items);
    }

    private DocumentResult ProcessSetBackgroundGradient(string filePath, PowerPointOperation op)
    {
        if (!op.SlideIndex.HasValue || string.IsNullOrWhiteSpace(op.GradientJson))
            return new DocumentResult(false, "Slide index and gradientJson are required");

        GradientFillOptions? gradient = ParseGradient(op.GradientJson);
        if (gradient == null)
            return new DocumentResult(false, "Invalid gradient JSON");

        return powerPointService.SetSlideBackgroundGradient(filePath, op.SlideIndex.Value, gradient);
    }

    private DocumentResult ProcessSetSlideSize(string filePath, PowerPointOperation op)
    {
        return powerPointService.SetSlideSize(filePath, op.SlideSize ?? "Widescreen");
    }

    // Helper methods for JSON parsing
    private GradientFillOptions? ParseGradient(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        try { return JsonSerializer.Deserialize<GradientFillOptions>(json, JsonOptions); }
        catch { return null; }
    }

    private RichParagraph[]? ParseParagraphs(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        try { return JsonSerializer.Deserialize<RichParagraph[]>(json, JsonOptions); }
        catch { return null; }
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
