using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeMCP.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WpTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace OfficeMCP.Services;

/// <summary>
/// Service for creating and manipulating Word documents using OpenXML.
/// </summary>
public sealed class WordDocumentService : IWordDocumentService
{
    private const int TwipsPerInch = 1440;
    private const int EmusPerInch = 914400;

    public DocumentResult CreateDocument(string filePath, string? title = null, PageLayoutOptions? layout = null)
    {
        try
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            if (layout != null)
            {
                ApplyPageLayout(mainPart.Document.Body!, layout);
            }

            if (!string.IsNullOrEmpty(title))
            {
                AddHeadingInternal(mainPart.Document.Body!, title, 1, null);
            }

            document.Save();
            return new DocumentResult(true, $"Document created successfully at {filePath}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to create document: {ex.Message}");
        }
    }

    public DocumentResult AddParagraph(string filePath, string text, TextFormatting? textFormat = null, ParagraphFormatting? paragraphFormat = null)
    {
        try
        {
            using var document = WordprocessingDocument.Open(filePath, true);
            var body = document.MainDocumentPart?.Document.Body;
            
            if (body == null)
                return new DocumentResult(false, "Document body not found");

            var paragraph = CreateParagraph(text, textFormat, paragraphFormat);
            
            // Insert before section properties if they exist
            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr != null)
            {
                body.InsertBefore(paragraph, sectPr);
            }
            else
            {
                body.AppendChild(paragraph);
            }

            document.Save();
            return new DocumentResult(true, "Paragraph added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add paragraph: {ex.Message}");
        }
    }

    public DocumentResult AddHeading(string filePath, string text, int level = 1, TextFormatting? textFormat = null)
    {
        try
        {
            using var document = WordprocessingDocument.Open(filePath, true);
            var body = document.MainDocumentPart?.Document.Body;
            
            if (body == null)
                return new DocumentResult(false, "Document body not found");

            AddHeadingInternal(body, text, level, textFormat);
            document.Save();
            return new DocumentResult(true, $"Heading level {level} added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add heading: {ex.Message}");
        }
    }

    public DocumentResult AddTable(string filePath, string[][] data, TableFormatting? tableFormat = null)
    {
        try
        {
            if (data.Length == 0)
                return new DocumentResult(false, "Table data cannot be empty");

            using var document = WordprocessingDocument.Open(filePath, true);
            var body = document.MainDocumentPart?.Document.Body;
            
            if (body == null)
                return new DocumentResult(false, "Document body not found");

            var table = CreateTable(data, tableFormat);
            
            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr != null)
            {
                body.InsertBefore(table, sectPr);
            }
            else
            {
                body.AppendChild(table);
            }

            document.Save();
            return new DocumentResult(true, $"Table with {data.Length} rows added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add table: {ex.Message}");
        }
    }

    public DocumentResult AddImage(string filePath, string imagePath, ImageOptions? options = null)
    {
        try
        {
            if (!File.Exists(imagePath))
                return new DocumentResult(false, $"Image file not found: {imagePath}");

            options ??= new ImageOptions();

            using var document = WordprocessingDocument.Open(filePath, true);
            var mainPart = document.MainDocumentPart;
            var body = mainPart?.Document.Body;
            
            if (body == null || mainPart == null)
                return new DocumentResult(false, "Document body not found");

            var imagePart = mainPart.AddImagePart(GetImagePartType(imagePath));
            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);
            var element = CreateImageElement(relationshipId, options);

            var paragraph = new Paragraph(new Run(element));
            
            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr != null)
            {
                body.InsertBefore(paragraph, sectPr);
            }
            else
            {
                body.AppendChild(paragraph);
            }

            document.Save();
            return new DocumentResult(true, "Image added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add image: {ex.Message}");
        }
    }

    public DocumentResult AddHeader(string filePath, HeaderFooterOptions options)
    {
        try
        {
            using var document = WordprocessingDocument.Open(filePath, true);
            var mainPart = document.MainDocumentPart;
            
            if (mainPart == null)
                return new DocumentResult(false, "Main document part not found");

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            var header = CreateHeaderContent(options);
            headerPart.Header = header;

            EnsureSectionProperties(mainPart.Document.Body!);
            var sectPr = mainPart.Document.Body!.Elements<SectionProperties>().First();
            
            sectPr.RemoveAllChildren<HeaderReference>();
            sectPr.PrependChild(new HeaderReference 
            { 
                Type = HeaderFooterValues.Default, 
                Id = mainPart.GetIdOfPart(headerPart) 
            });

            document.Save();
            return new DocumentResult(true, "Header added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add header: {ex.Message}");
        }
    }

    public DocumentResult AddFooter(string filePath, HeaderFooterOptions options)
    {
        try
        {
            using var document = WordprocessingDocument.Open(filePath, true);
            var mainPart = document.MainDocumentPart;
            
            if (mainPart == null)
                return new DocumentResult(false, "Main document part not found");

            var footerPart = mainPart.AddNewPart<FooterPart>();
            var footer = CreateFooterContent(options);
            footerPart.Footer = footer;

            EnsureSectionProperties(mainPart.Document.Body!);
            var sectPr = mainPart.Document.Body!.Elements<SectionProperties>().First();
            
            sectPr.RemoveAllChildren<FooterReference>();
            sectPr.PrependChild(new FooterReference 
            { 
                Type = HeaderFooterValues.Default, 
                Id = mainPart.GetIdOfPart(footerPart) 
            });

            document.Save();
            return new DocumentResult(true, "Footer added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add footer: {ex.Message}");
        }
    }

    public DocumentResult AddPageBreak(string filePath)
    {
        try
        {
            using var document = WordprocessingDocument.Open(filePath, true);
            var body = document.MainDocumentPart?.Document.Body;
            
            if (body == null)
                return new DocumentResult(false, "Document body not found");

            var paragraph = new Paragraph(
                new Run(
                    new Break { Type = BreakValues.Page }
                )
            );

            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr != null)
            {
                body.InsertBefore(paragraph, sectPr);
            }
            else
            {
                body.AppendChild(paragraph);
            }

            document.Save();
            return new DocumentResult(true, "Page break added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add page break: {ex.Message}");
        }
    }

    public DocumentResult AddBulletList(string filePath, string[] items, TextFormatting? textFormat = null)
    {
        return AddList(filePath, items, false, textFormat);
    }

    public DocumentResult AddNumberedList(string filePath, string[] items, TextFormatting? textFormat = null)
    {
        return AddList(filePath, items, true, textFormat);
    }

    public ContentResult GetDocumentText(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var document = WordprocessingDocument.Open(filePath, false);
            var body = document.MainDocumentPart?.Document.Body;
            
            if (body == null)
                return new ContentResult(false, null, "Document body not found");

            var paragraphs = body.Descendants<Paragraph>().ToList();
            var textContent = string.Join(Environment.NewLine, 
                paragraphs.Select(p => p.InnerText).Where(t => !string.IsNullOrWhiteSpace(t)));

            return new ContentResult(true, textContent, TotalParagraphs: paragraphs.Count);
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to read document: {ex.Message}");
        }
    }

    public ContentResult GetParagraphText(string filePath, int paragraphIndex)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var document = WordprocessingDocument.Open(filePath, false);
            var body = document.MainDocumentPart?.Document.Body;
            
            if (body == null)
                return new ContentResult(false, null, "Document body not found");

            var paragraphs = body.Descendants<Paragraph>().ToList();
            
            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
                return new ContentResult(false, null, $"Paragraph index {paragraphIndex} is out of range. Document has {paragraphs.Count} paragraphs.");

            return new ContentResult(true, paragraphs[paragraphIndex].InnerText, TotalParagraphs: paragraphs.Count);
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to read paragraph: {ex.Message}");
        }
    }

    public ContentResult GetParagraphRange(string filePath, int startIndex, int endIndex)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var document = WordprocessingDocument.Open(filePath, false);
            var body = document.MainDocumentPart?.Document.Body;
            
            if (body == null)
                return new ContentResult(false, null, "Document body not found");

            var paragraphs = body.Descendants<Paragraph>().ToList();
            
            if (startIndex < 0 || endIndex >= paragraphs.Count || startIndex > endIndex)
                return new ContentResult(false, null, $"Invalid range [{startIndex}, {endIndex}]. Document has {paragraphs.Count} paragraphs.");

            var textContent = string.Join(Environment.NewLine, 
                paragraphs.Skip(startIndex).Take(endIndex - startIndex + 1).Select(p => p.InnerText));

            return new ContentResult(true, textContent, TotalParagraphs: paragraphs.Count);
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to read paragraphs: {ex.Message}");
        }
    }

    public DocumentResult SetPageLayout(string filePath, PageLayoutOptions options)
    {
        try
        {
            using var document = WordprocessingDocument.Open(filePath, true);
            var body = document.MainDocumentPart?.Document.Body;
            
            if (body == null)
                return new DocumentResult(false, "Document body not found");

            ApplyPageLayout(body, options);
            document.Save();
            return new DocumentResult(true, "Page layout updated successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to set page layout: {ex.Message}");
        }
    }

    public DocumentResult AddMarkdownContent(string filePath, string markdown, string? baseImagePath = null)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(markdown))
                return new DocumentResult(false, "Markdown content cannot be empty");

            using var document = WordprocessingDocument.Open(filePath, true);
            var mainPart = document.MainDocumentPart;
            var body = mainPart?.Document.Body;
            
            if (body == null || mainPart == null)
                return new DocumentResult(false, "Document body not found");

            var elements = MarkdownParser.Parse(markdown);
            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();

            foreach (var element in elements)
            {
                var wordElements = ConvertMarkdownElement(element, mainPart, baseImagePath);
                foreach (var wordElement in wordElements)
                {
                    if (sectPr != null)
                    {
                        body.InsertBefore(wordElement, sectPr);
                    }
                    else
                    {
                        body.AppendChild(wordElement);
                    }
                }
            }

            document.Save();
            return new DocumentResult(true, $"Markdown content added successfully ({elements.Count} elements)", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add markdown content: {ex.Message}");
        }
    }

    #region Private Helper Methods

    private static Paragraph CreateParagraph(string text, TextFormatting? textFormat, ParagraphFormatting? paragraphFormat)
    {
        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        
        if (textFormat != null)
        {
            run.PrependChild(CreateRunProperties(textFormat));
        }

        var paragraph = new Paragraph(run);
        
        if (paragraphFormat != null)
        {
            paragraph.PrependChild(CreateParagraphProperties(paragraphFormat));
        }

        return paragraph;
    }

    private static void AddHeadingInternal(Body body, string text, int level, TextFormatting? textFormat)
    {
        level = Math.Clamp(level, 1, 9);
        
        var fontSize = level switch
        {
            1 => 32,
            2 => 26,
            3 => 22,
            4 => 20,
            5 => 18,
            _ => 16
        };

        var effectiveFormat = textFormat ?? new TextFormatting();
        effectiveFormat = effectiveFormat with 
        { 
            Bold = true, 
            FontSize = effectiveFormat.FontSize ?? fontSize 
        };

        var run = new Run(new Text(text));
        run.PrependChild(CreateRunProperties(effectiveFormat));

        var paragraph = new Paragraph(run);
        var pPr = new ParagraphProperties(
            new ParagraphStyleId { Val = $"Heading{level}" },
            new SpacingBetweenLines { Before = "240", After = "120" }
        );
        paragraph.PrependChild(pPr);

        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectPr != null)
        {
            body.InsertBefore(paragraph, sectPr);
        }
        else
        {
            body.AppendChild(paragraph);
        }
    }

    private static RunProperties CreateRunProperties(TextFormatting format)
    {
        var rPr = new RunProperties();

        if (format.Bold)
            rPr.AppendChild(new Bold());
        
        if (format.Italic)
            rPr.AppendChild(new Italic());
        
        if (format.Underline)
            rPr.AppendChild(new Underline { Val = UnderlineValues.Single });
        
        if (format.Strikethrough)
            rPr.AppendChild(new Strike());
        
        if (!string.IsNullOrEmpty(format.FontName))
            rPr.AppendChild(new RunFonts { Ascii = format.FontName, HighAnsi = format.FontName });
        
        if (format.FontSize.HasValue)
            rPr.AppendChild(new FontSize { Val = (format.FontSize.Value * 2).ToString() });
        
        if (!string.IsNullOrEmpty(format.FontColor))
            rPr.AppendChild(new Color { Val = format.FontColor.TrimStart('#') });
        
        if (!string.IsNullOrEmpty(format.HighlightColor) && 
            Enum.TryParse<HighlightColorValues>(format.HighlightColor, true, out var highlight))
            rPr.AppendChild(new Highlight { Val = highlight });

        return rPr;
    }

    private static ParagraphProperties CreateParagraphProperties(ParagraphFormatting format)
    {
        var pPr = new ParagraphProperties();

        var justification = format.Alignment.ToLowerInvariant() switch
        {
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            "justify" => JustificationValues.Both,
            _ => JustificationValues.Left
        };
        pPr.AppendChild(new Justification { Val = justification });

        var spacing = new SpacingBetweenLines();
        if (format.LineSpacing.HasValue)
            spacing.Line = ((int)(format.LineSpacing.Value * 240)).ToString();
        if (format.SpacingBefore.HasValue)
            spacing.Before = ((int)(format.SpacingBefore.Value * TwipsPerInch)).ToString();
        if (format.SpacingAfter.HasValue)
            spacing.After = ((int)(format.SpacingAfter.Value * TwipsPerInch)).ToString();
        pPr.AppendChild(spacing);

        if (format.FirstLineIndent.HasValue || format.LeftIndent.HasValue || format.RightIndent.HasValue)
        {
            var indentation = new Indentation();
            if (format.FirstLineIndent.HasValue)
                indentation.FirstLine = ((int)(format.FirstLineIndent.Value * TwipsPerInch)).ToString();
            if (format.LeftIndent.HasValue)
                indentation.Left = ((int)(format.LeftIndent.Value * TwipsPerInch)).ToString();
            if (format.RightIndent.HasValue)
                indentation.Right = ((int)(format.RightIndent.Value * TwipsPerInch)).ToString();
            pPr.AppendChild(indentation);
        }

        return pPr;
    }

    private static Table CreateTable(string[][] data, TableFormatting? format)
    {
        format ??= new TableFormatting();
        
        var table = new Table();
        
        // Table properties
        var tblPr = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = (uint)(format.BorderWidth * 4), Color = format.BorderColor?.TrimStart('#') ?? "000000" },
                new BottomBorder { Val = BorderValues.Single, Size = (uint)(format.BorderWidth * 4), Color = format.BorderColor?.TrimStart('#') ?? "000000" },
                new LeftBorder { Val = BorderValues.Single, Size = (uint)(format.BorderWidth * 4), Color = format.BorderColor?.TrimStart('#') ?? "000000" },
                new RightBorder { Val = BorderValues.Single, Size = (uint)(format.BorderWidth * 4), Color = format.BorderColor?.TrimStart('#') ?? "000000" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = (uint)(format.BorderWidth * 4), Color = format.BorderColor?.TrimStart('#') ?? "000000" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = (uint)(format.BorderWidth * 4), Color = format.BorderColor?.TrimStart('#') ?? "000000" }
            ),
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
        );
        table.AppendChild(tblPr);

        for (int rowIndex = 0; rowIndex < data.Length; rowIndex++)
        {
            var row = new TableRow();
            var rowData = data[rowIndex];
            var isHeader = format.HasHeader && rowIndex == 0;
            var isAlternate = !isHeader && format.AlternateRowColor != null && rowIndex % 2 == 1;

            foreach (var cellText in rowData)
            {
                var cell = new WpTableCell();
                
                var tcPr = new TableCellProperties(
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                );

                if (isHeader && !string.IsNullOrEmpty(format.HeaderBackgroundColor))
                {
                    tcPr.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = format.HeaderBackgroundColor.TrimStart('#') });
                }
                else if (isAlternate)
                {
                    tcPr.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = format.AlternateRowColor!.TrimStart('#') });
                }

                cell.AppendChild(tcPr);

                var run = new Run(new Text(cellText ?? string.Empty));
                if (isHeader)
                {
                    run.PrependChild(new RunProperties(new Bold()));
                }
                cell.AppendChild(new Paragraph(run));

                row.AppendChild(cell);
            }

            table.AppendChild(row);
        }

        return table;
    }

    private static PartTypeInfo GetImagePartType(string imagePath)
    {
        var extension = Path.GetExtension(imagePath).ToLowerInvariant();
        return extension switch
        {
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".png" => ImagePartType.Png,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tiff" or ".tif" => ImagePartType.Tiff,
            _ => ImagePartType.Png
        };
    }

    private static Drawing CreateImageElement(string relationshipId, ImageOptions options)
    {
        var element = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = options.WidthEmu, Cy = options.HeightEmu },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = 1U, Name = "Picture 1", Description = options.AltText ?? string.Empty },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "Image" },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = options.WidthEmu, Cy = options.HeightEmu }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            }
        );

        return element;
    }

    private static Header CreateHeaderContent(HeaderFooterOptions options)
    {
        var header = new Header();
        var paragraph = CreateHeaderFooterParagraph(options);
        header.AppendChild(paragraph);
        return header;
    }

    private static Footer CreateFooterContent(HeaderFooterOptions options)
    {
        var footer = new Footer();
        var paragraph = CreateHeaderFooterParagraph(options);
        footer.AppendChild(paragraph);
        return footer;
    }

    private static Paragraph CreateHeaderFooterParagraph(HeaderFooterOptions options)
    {
        var paragraph = new Paragraph();
        var pPr = new ParagraphProperties();

        if (options.CenterContent != null || options.IncludePageNumber)
        {
            pPr.AppendChild(new Justification { Val = JustificationValues.Center });
        }

        paragraph.AppendChild(pPr);

        if (!string.IsNullOrEmpty(options.LeftContent))
        {
            paragraph.AppendChild(new Run(new Text(options.LeftContent) { Space = SpaceProcessingModeValues.Preserve }));
        }

        if (!string.IsNullOrEmpty(options.CenterContent))
        {
            paragraph.AppendChild(new Run(new TabChar()));
            paragraph.AppendChild(new Run(new Text(options.CenterContent) { Space = SpaceProcessingModeValues.Preserve }));
        }

        if (options.IncludePageNumber)
        {
            paragraph.AppendChild(new Run(new TabChar()));
            paragraph.AppendChild(new Run(new Text("Page ") { Space = SpaceProcessingModeValues.Preserve }));
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            paragraph.AppendChild(new Run(new FieldCode(" PAGE ")));
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        if (options.IncludeDate)
        {
            paragraph.AppendChild(new Run(new TabChar()));
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            paragraph.AppendChild(new Run(new FieldCode(" DATE \\@ \"MM/dd/yyyy\" ")));
            paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        if (!string.IsNullOrEmpty(options.RightContent))
        {
            paragraph.AppendChild(new Run(new TabChar()));
            paragraph.AppendChild(new Run(new Text(options.RightContent) { Space = SpaceProcessingModeValues.Preserve }));
        }

        return paragraph;
    }

    private static void EnsureSectionProperties(Body body)
    {
        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectPr == null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }
    }

    private static void ApplyPageLayout(Body body, PageLayoutOptions options)
    {
        EnsureSectionProperties(body);
        var sectPr = body.Elements<SectionProperties>().First();

        // Page size
        var pageSize = sectPr.Elements<PageSize>().FirstOrDefault();
        if (pageSize == null)
        {
            pageSize = new PageSize();
            sectPr.PrependChild(pageSize);
        }

        var (width, height) = options.PageSize.ToUpperInvariant() switch
        {
            "LETTER" => (12240U, 15840U),
            "LEGAL" => (12240U, 20160U),
            "A4" => (11906U, 16838U),
            "A3" => (16838U, 23811U),
            _ => (12240U, 15840U)
        };

        if (options.Orientation.Equals("Landscape", StringComparison.OrdinalIgnoreCase))
        {
            pageSize.Width = height;
            pageSize.Height = width;
            pageSize.Orient = PageOrientationValues.Landscape;
        }
        else
        {
            pageSize.Width = width;
            pageSize.Height = height;
            pageSize.Orient = PageOrientationValues.Portrait;
        }

        // Margins
        var pageMargin = sectPr.Elements<PageMargin>().FirstOrDefault();
        if (pageMargin == null)
        {
            pageMargin = new PageMargin();
            sectPr.AppendChild(pageMargin);
        }

        pageMargin.Top = (int)((options.MarginTop ?? 1.0) * TwipsPerInch);
        pageMargin.Bottom = (int)((options.MarginBottom ?? 1.0) * TwipsPerInch);
        pageMargin.Left = (uint)((options.MarginLeft ?? 1.0) * TwipsPerInch);
        pageMargin.Right = (uint)((options.MarginRight ?? 1.0) * TwipsPerInch);
    }

    private DocumentResult AddList(string filePath, string[] items, bool numbered, TextFormatting? textFormat)
    {
        try
        {
            if (items.Length == 0)
                return new DocumentResult(false, "List items cannot be empty");

            using var document = WordprocessingDocument.Open(filePath, true);
            var mainPart = document.MainDocumentPart;
            var body = mainPart?.Document.Body;
            
            if (body == null || mainPart == null)
                return new DocumentResult(false, "Document body not found");

            // Ensure numbering definitions exist
            var numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering();
            }

            var numbering = numberingPart.Numbering;
            var abstractNumId = numbering.Elements<AbstractNum>().Count() + 1;
            var numId = numbering.Elements<NumberingInstance>().Count() + 1;

            // Create abstract numbering
            var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };
            var level = new Level { LevelIndex = 0 };
            
            if (numbered)
            {
                level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Decimal });
                level.AppendChild(new LevelText { Val = "%1." });
            }
            else
            {
                level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Bullet });
                level.AppendChild(new LevelText { Val = "•" });
            }
            
            level.AppendChild(new StartNumberingValue { Val = 1 });
            level.AppendChild(new ParagraphProperties(
                new Indentation { Left = "720", Hanging = "360" }
            ));
            
            abstractNum.AppendChild(level);
            numbering.InsertAt(abstractNum, 0);

            // Create numbering instance
            var numberingInstance = new NumberingInstance { NumberID = numId };
            numberingInstance.AppendChild(new AbstractNumId { Val = abstractNumId });
            numbering.AppendChild(numberingInstance);

            // Add list items
            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
            
            foreach (var item in items)
            {
                var run = new Run(new Text(item));
                if (textFormat != null)
                {
                    run.PrependChild(CreateRunProperties(textFormat));
                }

                var paragraph = new Paragraph(run);
                var pPr = new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = numId }
                    )
                );
                paragraph.PrependChild(pPr);

                if (sectPr != null)
                {
                    body.InsertBefore(paragraph, sectPr);
                }
                else
                {
                    body.AppendChild(paragraph);
                }
            }

            document.Save();
            return new DocumentResult(true, $"{(numbered ? "Numbered" : "Bullet")} list with {items.Length} items added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add list: {ex.Message}");
        }
    }

    #region Markdown Conversion Methods

    private static List<OpenXmlElement> ConvertMarkdownElement(MarkdownElement element, MainDocumentPart mainPart, string? baseImagePath)
    {
        return element switch
        {
            MarkdownHeading heading => [CreateMarkdownHeading(heading)],
            MarkdownParagraph para => [CreateMarkdownParagraph(para.Inlines)],
            MarkdownBulletList bulletList => CreateMarkdownList(bulletList.Items, false, mainPart),
            MarkdownNumberedList numberedList => CreateMarkdownList(numberedList.Items, true, mainPart),
            MarkdownCodeBlock codeBlock => [CreateCodeBlockParagraph(codeBlock)],
            MarkdownBlockquote blockquote => [CreateBlockquoteParagraph(blockquote)],
            MarkdownHorizontalRule => [CreateHorizontalRule()],
            MarkdownTable table => [CreateMarkdownTable(table)],
            MarkdownImage image => CreateMarkdownImage(image, mainPart, baseImagePath),
            _ => []
        };
    }

    private static Paragraph CreateMarkdownHeading(MarkdownHeading heading)
    {
        var fontSize = heading.Level switch
        {
            1 => 32,
            2 => 26,
            3 => 22,
            4 => 20,
            5 => 18,
            _ => 16
        };

        var paragraph = new Paragraph();
        var pPr = new ParagraphProperties(
            new ParagraphStyleId { Val = $"Heading{heading.Level}" },
            new SpacingBetweenLines { Before = "240", After = "120" }
        );
        paragraph.AppendChild(pPr);

        foreach (var inline in heading.Inlines)
        {
            var run = CreateRunFromInline(inline);
            // Ensure heading text is bold
            var rPr = run.GetFirstChild<RunProperties>() ?? new RunProperties();
            if (rPr.Bold == null)
            {
                rPr.PrependChild(new Bold());
            }
            if (rPr.FontSize == null)
            {
                rPr.AppendChild(new FontSize { Val = (fontSize * 2).ToString() });
            }
            if (run.GetFirstChild<RunProperties>() == null)
            {
                run.PrependChild(rPr);
            }
            paragraph.AppendChild(run);
        }

        return paragraph;
    }

    private static Paragraph CreateMarkdownParagraph(List<MarkdownInline> inlines)
    {
        var paragraph = new Paragraph();
        
        foreach (var inline in inlines)
        {
            var run = CreateRunFromInline(inline);
            paragraph.AppendChild(run);
        }

        return paragraph;
    }

    private static Run CreateRunFromInline(MarkdownInline inline)
    {
        return inline switch
        {
            MarkdownText text => new Run(new Text(text.Text) { Space = SpaceProcessingModeValues.Preserve }),
            MarkdownBold bold => CreateFormattedRun(bold.Text, new RunProperties(new Bold())),
            MarkdownItalic italic => CreateFormattedRun(italic.Text, new RunProperties(new Italic())),
            MarkdownBoldItalic boldItalic => CreateFormattedRun(boldItalic.Text, new RunProperties(new Bold(), new Italic())),
            MarkdownStrikethrough strike => CreateFormattedRun(strike.Text, new RunProperties(new Strike())),
            MarkdownCode code => CreateCodeRun(code.Text),
            MarkdownLink link => CreateHyperlinkRun(link),
            MarkdownInlineImage image => new Run(new Text($"[Image: {image.AltText}]") { Space = SpaceProcessingModeValues.Preserve }),
            _ => new Run()
        };
    }

    private static Run CreateFormattedRun(string text, RunProperties runProperties)
    {
        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        run.PrependChild(runProperties);
        return run;
    }

    private static Run CreateCodeRun(string text)
    {
        var rPr = new RunProperties(
            new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" },
            new Shading { Val = ShadingPatternValues.Clear, Fill = "E8E8E8" },
            new FontSize { Val = "20" }
        );
        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        run.PrependChild(rPr);
        return run;
    }

    private static Run CreateHyperlinkRun(MarkdownLink link)
    {
        var rPr = new RunProperties(
            new Color { Val = "0563C1" },
            new Underline { Val = UnderlineValues.Single }
        );
        var run = new Run(new Text(link.Text) { Space = SpaceProcessingModeValues.Preserve });
        run.PrependChild(rPr);
        // Note: Full hyperlink support would require adding relationship to mainPart
        return run;
    }

    private static List<OpenXmlElement> CreateMarkdownList(List<MarkdownListItem> items, bool numbered, MainDocumentPart mainPart)
    {
        var elements = new List<OpenXmlElement>();
        
        // Ensure numbering definitions exist
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
        }

        var numbering = numberingPart.Numbering;
        var abstractNumId = numbering.Elements<AbstractNum>().Count() + 1;
        var numId = numbering.Elements<NumberingInstance>().Count() + 1;

        // Create abstract numbering
        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };
        var level = new Level { LevelIndex = 0 };
        
        if (numbered)
        {
            level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Decimal });
            level.AppendChild(new LevelText { Val = "%1." });
        }
        else
        {
            level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Bullet });
            level.AppendChild(new LevelText { Val = "•" });
        }
        
        level.AppendChild(new StartNumberingValue { Val = 1 });
        level.AppendChild(new ParagraphProperties(
            new Indentation { Left = "720", Hanging = "360" }
        ));
        
        abstractNum.AppendChild(level);
        numbering.InsertAt(abstractNum, 0);

        // Create numbering instance
        var numberingInstance = new NumberingInstance { NumberID = numId };
        numberingInstance.AppendChild(new AbstractNumId { Val = abstractNumId });
        numbering.AppendChild(numberingInstance);

        // Create list item paragraphs
        foreach (var item in items)
        {
            var paragraph = new Paragraph();
            var pPr = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = numId }
                )
            );
            paragraph.AppendChild(pPr);

            foreach (var inline in item.Inlines)
            {
                paragraph.AppendChild(CreateRunFromInline(inline));
            }

            elements.Add(paragraph);
        }

        return elements;
    }

    private static Paragraph CreateCodeBlockParagraph(MarkdownCodeBlock codeBlock)
    {
        var paragraph = new Paragraph();
        
        var pPr = new ParagraphProperties(
            new Shading { Val = ShadingPatternValues.Clear, Fill = "F5F5F5" },
            new SpacingBetweenLines { Before = "120", After = "120" },
            new Indentation { Left = "720" }
        );
        paragraph.AppendChild(pPr);

        var lines = codeBlock.Code.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            var rPr = new RunProperties(
                new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" },
                new FontSize { Val = "20" }
            );
            var run = new Run(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
            run.PrependChild(rPr);
            paragraph.AppendChild(run);

            if (i < lines.Length - 1)
            {
                paragraph.AppendChild(new Run(new Break()));
            }
        }

        return paragraph;
    }

    private static Paragraph CreateBlockquoteParagraph(MarkdownBlockquote blockquote)
    {
        var paragraph = new Paragraph();
        
        var pPr = new ParagraphProperties(
            new Indentation { Left = "720" },
            new ParagraphBorders(
                new LeftBorder { Val = BorderValues.Single, Size = 24, Color = "808080", Space = 4 }
            ),
            new SpacingBetweenLines { Before = "120", After = "120" }
        );
        paragraph.AppendChild(pPr);

        foreach (var inline in blockquote.Inlines)
        {
            var run = CreateRunFromInline(inline);
            var rPr = run.GetFirstChild<RunProperties>() ?? new RunProperties();
            rPr.AppendChild(new Italic());
            rPr.AppendChild(new Color { Val = "666666" });
            if (run.GetFirstChild<RunProperties>() == null)
            {
                run.PrependChild(rPr);
            }
            paragraph.AppendChild(run);
        }

        return paragraph;
    }

    private static Paragraph CreateHorizontalRule()
    {
        var paragraph = new Paragraph();
        
        var pPr = new ParagraphProperties(
            new ParagraphBorders(
                new BottomBorder { Val = BorderValues.Single, Size = 6, Color = "000000", Space = 1 }
            ),
            new SpacingBetweenLines { Before = "240", After = "240" }
        );
        paragraph.AppendChild(pPr);

        return paragraph;
    }

    private static Table CreateMarkdownTable(MarkdownTable mdTable)
    {
        var table = new Table();
        
        var tblPr = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" }
            ),
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
        );
        table.AppendChild(tblPr);

        // Header row
        var headerRow = new TableRow();
        foreach (var header in mdTable.Headers)
        {
            var cell = new WpTableCell();
            var tcPr = new TableCellProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = "E0E0E0" }
            );
            cell.AppendChild(tcPr);
            
            var run = new Run(new Text(header));
            run.PrependChild(new RunProperties(new Bold()));
            cell.AppendChild(new Paragraph(run));
            headerRow.AppendChild(cell);
        }
        table.AppendChild(headerRow);

        // Data rows
        foreach (var rowData in mdTable.Rows)
        {
            var row = new TableRow();
            foreach (var cellText in rowData)
            {
                var cell = new WpTableCell();
                cell.AppendChild(new Paragraph(new Run(new Text(cellText))));
                row.AppendChild(cell);
            }
            table.AppendChild(row);
        }

        return table;
    }

    private static List<OpenXmlElement> CreateMarkdownImage(MarkdownImage image, MainDocumentPart mainPart, string? baseImagePath)
    {
        var imagePath = image.Url;
        
        // Handle relative paths
        if (!Path.IsPathRooted(imagePath) && !string.IsNullOrEmpty(baseImagePath))
        {
            imagePath = Path.Combine(baseImagePath, imagePath);
        }

        // If image doesn't exist, create a placeholder paragraph
        if (!File.Exists(imagePath))
        {
            var placeholder = new Paragraph(
                new Run(new Text($"[Image not found: {image.AltText ?? image.Url}]") { Space = SpaceProcessingModeValues.Preserve })
            );
            return [placeholder];
        }

        try
        {
            var imagePart = mainPart.AddImagePart(GetImagePartType(imagePath));
            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);
            var options = new ImageOptions(
                WidthEmu: 4 * EmusPerInch,
                HeightEmu: 3 * EmusPerInch,
                AltText: image.AltText
            );
            var drawing = CreateImageElement(relationshipId, options);
            var paragraph = new Paragraph(new Run(drawing));
            
            return [paragraph];
        }
        catch
        {
            var errorPara = new Paragraph(
                new Run(new Text($"[Error loading image: {image.AltText ?? image.Url}]") { Space = SpaceProcessingModeValues.Preserve })
            );
            return [errorPara];
        }
    }

    #endregion

    #endregion
}
