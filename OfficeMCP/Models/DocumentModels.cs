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
/// Result of a document operation.
/// </summary>
public record DocumentResult(
    bool Success,
    string Message,
    string? FilePath = null,
    string? AdditionalInfo = null
);

/// <summary>
/// Result of a content extraction operation.
/// </summary>
public record ContentResult(
    bool Success,
    string? Content,
    string? ErrorMessage = null,
    int? TotalParagraphs = null,
    int? TotalPages = null
);
