namespace OfficeMCP.Models;

/// <summary>
/// Represents text formatting options for document content.
/// </summary>
public record TextFormatting(
    bool Bold = false,
    bool Italic = false,
    bool Underline = false,
    bool Strikethrough = false,
    string? FontName = null,
    int? FontSize = null,
    string? FontColor = null,
    string? HighlightColor = null
);

/// <summary>
/// Represents paragraph formatting options.
/// </summary>
public record ParagraphFormatting(
    string Alignment = "Left",
    double? LineSpacing = null,
    double? SpacingBefore = null,
    double? SpacingAfter = null,
    double? FirstLineIndent = null,
    double? LeftIndent = null,
    double? RightIndent = null
);

/// <summary>
/// Represents a table cell with content and formatting.
/// </summary>
public record TableCell(
    string Content,
    TextFormatting? TextFormat = null,
    string? BackgroundColor = null,
    string HorizontalAlignment = "Left",
    string VerticalAlignment = "Center",
    int ColumnSpan = 1,
    int RowSpan = 1
);

/// <summary>
/// Represents table formatting options.
/// </summary>
public record TableFormatting(
    string? BorderColor = null,
    double BorderWidth = 1.0,
    bool HasHeader = true,
    string? HeaderBackgroundColor = null,
    string? AlternateRowColor = null
);

/// <summary>
/// Represents image positioning and sizing options.
/// </summary>
public record ImageOptions(
    long WidthEmu = 914400,
    long HeightEmu = 914400,
    string? AltText = null,
    string Positioning = "Inline",
    /// <summary>
    /// Shape to crop/clip the image into: Rectangle (default), Ellipse, RoundRectangle, Triangle, etc.
    /// Use "Ellipse" for circular avatar photos.
    /// </summary>
    string CropShape = "Rectangle",
    /// <summary>
    /// Rotation in degrees (0-360).
    /// </summary>
    double Rotation = 0.0,
    /// <summary>
    /// Border/outline color as hex (e.g., "4472C4"). No border if null.
    /// </summary>
    string? BorderColor = null,
    /// <summary>
    /// Border width in points.
    /// </summary>
    double BorderWidth = 0.0,
    /// <summary>
    /// Add drop shadow effect.
    /// </summary>
    bool HasShadow = false,
    /// <summary>
    /// 3D rotation angle around the Y axis in degrees (perspective tilt).
    /// Positive = rotate right edge away, negative = rotate left edge away.
    /// Use 15-30 degrees for subtle perspective card effects.
    /// </summary>
    double Perspective3DAngleY = 0.0,
    /// <summary>
    /// 3D rotation angle around the X axis in degrees (perspective tilt).
    /// Positive = rotate top edge away, negative = rotate bottom edge away.
    /// </summary>
    double Perspective3DAngleX = 0.0
);

/// <summary>
/// Represents header/footer content options.
/// </summary>
public record HeaderFooterOptions(
    string? LeftContent = null,
    string? CenterContent = null,
    string? RightContent = null,
    bool IncludePageNumber = false,
    bool IncludeDate = false
);

/// <summary>
/// Represents page layout options for Word documents.
/// </summary>
public record PageLayoutOptions(
    string Orientation = "Portrait",
    double? MarginTop = null,
    double? MarginBottom = null,
    double? MarginLeft = null,
    double? MarginRight = null,
    string PageSize = "Letter"
);

/// <summary>
/// Represents Excel cell formatting options.
/// </summary>
public record ExcelCellFormatting(
    bool Bold = false,
    bool Italic = false,
    string? FontColor = null,
    string? BackgroundColor = null,
    string? NumberFormat = null,
    string HorizontalAlignment = "General",
    string VerticalAlignment = "Bottom",
    bool WrapText = false,
    string? BorderStyle = null
);

/// <summary>
/// Represents a range of cells in Excel.
/// </summary>
public record CellRange(
    string StartCell,
    string EndCell
);

/// <summary>
/// Represents PowerPoint slide layout options.
/// </summary>
public record SlideLayoutOptions(
    string LayoutType = "Blank",
    string? BackgroundColor = null,
    GradientFillOptions? GradientBackground = null
);

/// <summary>
/// Represents a single text run with its own formatting within a paragraph.
/// Allows multiple differently-formatted segments within one text box.
/// </summary>
public record TextRun(
    string Text,
    TextFormatting? Format = null
);

/// <summary>
/// Represents a single paragraph with optional formatting and multiple text runs.
/// </summary>
public record RichParagraph(
    TextRun[]? Runs = null,
    string? Text = null,
    string Alignment = "Left",
    double? SpacingBeforePt = null,
    double? SpacingAfterPt = null,
    double? LineSpacingPercent = null,
    bool IsBullet = false,
    string? BulletChar = null,
    int? IndentLevel = null
);

/// <summary>
/// Gradient stop with position (0-100) and color.
/// </summary>
public record GradientStop(
    int Position,
    string Color,
    int? TransparencyPercent = null
);

/// <summary>
/// Gradient fill options for shapes and backgrounds.
/// </summary>
public record GradientFillOptions(
    GradientStop[] Stops,
    double Angle = 0.0,
    string GradientType = "Linear"
);

/// <summary>
/// Line/connector options for PowerPoint.
/// </summary>
public record LineOptions(
    long X1,
    long Y1,
    long X2,
    long Y2,
    string? LineColor = null,
    double LineWidth = 1.0,
    string DashStyle = "Solid",
    string? StartArrow = null,
    string? EndArrow = null
);

/// <summary>
/// Connector options for connecting two shapes.
/// </summary>
public record ConnectorOptions(
    long X1,
    long Y1,
    long X2,
    long Y2,
    string ConnectorType = "Straight",
    string? LineColor = null,
    double LineWidth = 1.0,
    string DashStyle = "Solid",
    string? StartArrow = null,
    string? EndArrow = null
);

/// <summary>
/// An item within a group shape.
/// </summary>
public record GroupShapeItem(
    string ItemType,
    long X,
    long Y,
    long Width,
    long Height,
    string? Text = null,
    TextFormatting? TextFormat = null,
    string? ShapeType = null,
    string? FillColor = null,
    string? BorderColor = null,
    double BorderWidth = 1.0,
    string? ImagePath = null
);

/// <summary>
/// Represents text box options for PowerPoint.
/// </summary>
public record TextBoxOptions(
    long X,
    long Y,
    long Width,
    long Height,
    string? BackgroundColor = null,
    string? BorderColor = null,
    TextFormatting? TextFormat = null,
    double BorderWidth = 1.0,
    string VerticalAlignment = "Top",
    double MarginLeftInches = 0.1,
    double MarginRightInches = 0.1,
    double MarginTopInches = 0.05,
    double MarginBottomInches = 0.05,
    bool WordWrap = true,
    string AutoFit = "None",
    double Rotation = 0.0,
    string Alignment = "Left",
    RichParagraph[]? Paragraphs = null,
    GradientFillOptions? GradientFill = null
);

/// <summary>
/// Represents shape options for PowerPoint.
/// </summary>
public record ShapeOptions(
    string ShapeType,
    long X,
    long Y,
    long Width,
    long Height,
    string? FillColor = null,
    string? BorderColor = null,
    double BorderWidth = 1.0,
    string? Text = null,
    TextFormatting? TextFormat = null,
    string TextAlignment = "Center",
    string VerticalTextAlignment = "Middle",
    double Rotation = 0.0,
    int? CornerRadiusPt = null,
    bool HasShadow = false,
    string DashStyle = "Solid",
    int TransparencyPercent = 0,
    GradientFillOptions? GradientFill = null,
    RichParagraph[]? Paragraphs = null,
    double MarginLeftInches = 0.1,
    double MarginRightInches = 0.1,
    double MarginTopInches = 0.05,
    double MarginBottomInches = 0.05,
    bool NoFill = false,
    /// <summary>
    /// 3D rotation angle around the Y axis in degrees (perspective tilt).
    /// </summary>
    double Perspective3DAngleY = 0.0,
    /// <summary>
    /// 3D rotation angle around the X axis in degrees (perspective tilt).
    /// </summary>
    double Perspective3DAngleX = 0.0
);

/// <summary>
/// Result of a document operation with actionable error information.
/// </summary>
public record DocumentResult(
    bool Success,
    string Message,
    string? FilePath = null,
    string? Format = null,
    string? Suggestion = null
);

/// <summary>
/// Result of a content extraction operation.
/// </summary>
public record ContentResult(
    bool Success,
    string? Content,
    string? ErrorMessage = null,
    int? TotalParagraphs = null,
    int? TotalPages = null,
    string? Format = null,
    string? Suggestion = null
);

/// <summary>
/// A single content item returned by GetRichContent: one of "heading", "paragraph", "image", or "table".
/// Items are returned in document reading order so that section headings and surrounding paragraphs
/// provide natural context for any inline images.
/// </summary>
public record DocumentContentItem(
    string Type,          // "heading" | "paragraph" | "image" | "table"
    string? Text = null,  // Present for heading/paragraph/table
    int? Level = null,    // Heading level 1-6 (heading only)
    string? Style = null, // Word paragraph style name (e.g., "Title", "Subtitle", "Heading1", "Normal", "Quote")
    string? MimeType = null,      // image/* (image only)
    string? ImageBase64 = null,   // Base64 image bytes (image only)
    string? AltText = null,       // Alt text stored in document (image only)
    int? WidthPx = null,
    int? HeightPx = null
);

// ============================================
// Batch Operation Models for AI Optimization
// ============================================

/// <summary>
/// Represents a single Word document operation for batch processing.
/// </summary>
public record WordOperation(
    string Type,
    string? Text = null,
    int? Level = null,
    string[]? Items = null,
    string[][]? TableData = null,
    string? ImagePath = null,
    double? WidthInches = null,
    double? HeightInches = null,
    string? AltText = null,
    bool? Bold = null,
    bool? Italic = null,
    bool? Underline = null,
    string? FontName = null,
    int? FontSize = null,
    string? FontColor = null,
    string? Alignment = null,
    double? LineSpacing = null,
    bool? HasHeader = null,
    string? HeaderBackgroundColor = null,
    string? AlternateRowColor = null,
    string? BorderColor = null,
    double? BorderWidth = null,
    string? LeftContent = null,
    string? CenterContent = null,
    string? RightContent = null,
    bool? IncludePageNumber = null,
    bool? IncludeDate = null,
    string? Markdown = null,
    string? BaseImagePath = null
);

/// <summary>
/// Represents a single Excel workbook operation for batch processing.
/// </summary>
public record ExcelOperation(
    string Type,
    string? SheetName = null,
    string? CellReference = null,
    string? StartCell = null,
    string? EndCell = null,
    string? Value = null,
    string[][]? Values = null,
    string[][]? TableData = null,
    bool? HasHeaders = null,
    string? Formula = null,
    string? ImagePath = null,
    double? WidthInches = null,
    double? HeightInches = null,
    string? AltText = null,
    int? ColumnIndex = null,
    int? RowIndex = null,
    double? Width = null,
    double? Height = null,
    bool? Bold = null,
    bool? Italic = null,
    bool? WrapText = null,
    string? NewSheetName = null,
    // Additional formatting properties
    string? FontColor = null,
    string? FillColor = null,
    string? NumberFormat = null,
    string? HorizontalAlignment = null,
    string? VerticalAlignment = null,
    string? BorderStyle = null
);

/// <summary>
/// Represents a single PowerPoint operation for batch processing.
/// </summary>
public record PowerPointOperation(
    string Type,
    int? SlideIndex = null,
    string? Title = null,
    string? Subtitle = null,
    string? Text = null,
    string[]? Points = null,
    string[][]? TableData = null,
    string? ImagePath = null,
    string? ShapeType = null,
    double? XInches = null,
    double? YInches = null,
    double? WidthInches = null,
    double? HeightInches = null,
    string? BackgroundColor = null,
    string? FillColor = null,
    string? BorderColor = null,
    double? BorderWidth = null,
    bool? Bold = null,
    bool? Italic = null,
    bool? Underline = null,
    int? FontSize = null,
    string? FontColor = null,
    string? FontName = null,
    string? AltText = null,
    string? Notes = null,
    int? SourceIndex = null,
    int? FromIndex = null,
    int? ToIndex = null,
    // Rich text support
    string? Alignment = null,
    string? VerticalAlignment = null,
    string? ParagraphsJson = null,
    // Shape enhancements
    double? Rotation = null,
    int? CornerRadiusPt = null,
    bool? HasShadow = null,
    string? DashStyle = null,
    int? TransparencyPercent = null,
    string? GradientJson = null,
    bool? NoFill = null,
    bool? WordWrap = null,
    string? AutoFit = null,
    // Line/connector support
    double? X2Inches = null,
    double? Y2Inches = null,
    string? StartArrow = null,
    string? EndArrow = null,
    string? ConnectorType = null,
    string? LineColor = null,
    double? LineWidth = null,
    // Group shape support
    string? GroupItemsJson = null,
    // Margin support
    double? MarginLeftInches = null,
    double? MarginRightInches = null,
    double? MarginTopInches = null,
    double? MarginBottomInches = null,
    // Slide size
    string? SlideSize = null,
    // Image crop shape (Ellipse for circular avatars)
    string? CropShape = null,
    // 3D perspective rotation
    double? Perspective3DAngleY = null,
    double? Perspective3DAngleX = null,
    // Base64 image data (alternative to ImagePath)
    string? ImageBase64 = null,
    string? ImageMimeType = null,
    // Z-order placement
    int? ZOrderPosition = null
);

/// <summary>
/// Result of a batch operation containing results for each operation.
/// </summary>
public record BatchOperationResult(
    bool Success,
    string Message,
    int TotalOperations,
    int SuccessfulOperations,
    int FailedOperations,
    List<OperationOutcome> Details
);

/// <summary>
/// Outcome of a single operation within a batch.
/// </summary>
public record OperationOutcome(
    int Index,
    string OperationType,
    bool Success,
    string Message
);

// ============================================
// Excel Formatting Models
// ============================================

/// <summary>
/// Result of getting cell/range formatting information.
/// </summary>
public record ExcelRangeFormattingResult(
    bool Success,
    string? ErrorMessage,
    List<ExcelCellInfo>? Cells,
    string? SheetName,
    string? Range
);

/// <summary>
/// Information about a single Excel cell including value and formatting.
/// </summary>
public record ExcelCellInfo(
    string CellReference,
    string? Value,
    string? Formula,
    ExcelCellFormattingInfo? Formatting
);

/// <summary>
/// Detailed formatting information for an Excel cell.
/// </summary>
public record ExcelCellFormattingInfo(
    bool Bold,
    bool Italic,
    bool Underline,
    string? FontName,
    double? FontSize,
    string? FontColor,
    string? BackgroundColor,
    string? NumberFormat,
    string HorizontalAlignment,
    string VerticalAlignment,
    bool WrapText,
    bool HasBorder,
    string? BorderStyle
);

/// <summary>
/// An image extracted from a document, with alt text, surrounding paragraph context,
/// and base64-encoded image data for AI vision analysis (OCR, captioning).
/// </summary>
public record ImageExtractionResult(
    int Index,
    string MimeType,
    string ImageBase64,
    string AltText,
    string ContextBefore,
    string ContextAfter,
    int? WidthPx = null,
    int? HeightPx = null,
    int? PageOrSlideNumber = null
);
