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
    string Positioning = "Inline"
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
    string? BackgroundColor = null
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
    TextFormatting? TextFormat = null
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
    double BorderWidth = 1.0
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
    int? FontSize = null,
    string? FontColor = null,
    string? AltText = null,
    string? Notes = null,
    int? SourceIndex = null,
    int? FromIndex = null,
    int? ToIndex = null
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
