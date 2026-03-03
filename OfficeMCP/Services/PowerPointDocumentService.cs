using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeMCP.Models;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeMCP.Services;

/// <summary>
/// Service for creating and manipulating PowerPoint presentations using OpenXML.
/// </summary>
public sealed class PowerPointDocumentService : IPowerPointDocumentService
{
    private const long EmusPerInch = 914400;
    private const long DefaultSlideWidth = 9144000;
    private const long DefaultSlideHeight = 6858000;
    private static uint _nextShapeId = 10000;

    /// <summary>
    /// Thread-safe shape ID generator to avoid duplicate IDs that corrupt PPTX files.
    /// Using new Random() per call was producing duplicates when shapes were added in quick succession.
    /// </summary>
    private static uint NextShapeId() => Interlocked.Increment(ref _nextShapeId);

    public DocumentResult CreatePresentation(string filePath, string? title = null)
    {
        try
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Create and save the base presentation structure, then dispose before opening again
            using (var presentation = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
            {
                var presentationPart = presentation.AddPresentationPart();
                presentationPart.Presentation = new Presentation(
                    new SlideIdList(),
                    new SlideSize { Cx = (int)DefaultSlideWidth, Cy = (int)DefaultSlideHeight },
                    new NotesSize { Cx = (int)DefaultSlideHeight, Cy = (int)DefaultSlideWidth }
                );

                // Add slide master
                var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                slideMasterPart.SlideMaster = CreateSlideMaster();

                var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
                slideLayoutPart.SlideLayout = CreateSlideLayout();

                slideMasterPart.SlideMaster.Append(new SlideLayoutIdList(
                    new SlideLayoutId { Id = 2147483649U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart) }
                ));

                presentationPart.Presentation.Append(new SlideMasterIdList(
                    new SlideMasterId { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) }
                ));

                // Add theme
                var themePart = slideMasterPart.AddNewPart<ThemePart>();
                themePart.Theme = CreateDefaultTheme();

                presentation.Save();
            } // File is fully closed/flushed here

            // Add first slide with title if provided (file is now closed so we can reopen it)
            if (!string.IsNullOrEmpty(title))
            {
                AddSlide(filePath);
                AddTitle(filePath, 0, title);
            }

            return new DocumentResult(true, $"Presentation created successfully at {filePath}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to create presentation: {ex.Message}");
        }
    }

    public DocumentResult AddSlide(string filePath, SlideLayoutOptions? layoutOptions = null)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var presentationPart = presentation.PresentationPart;
            
            if (presentationPart == null)
                return new DocumentResult(false, "Presentation part not found");

            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = CreateSlide(layoutOptions);

            // Link to slide layout
            var slideMasterPart = presentationPart.SlideMasterParts.FirstOrDefault();
            var slideLayoutPart = slideMasterPart?.SlideLayoutParts.FirstOrDefault();
            
            if (slideLayoutPart != null)
            {
                slidePart.AddPart(slideLayoutPart);
            }

            var slideIdList = presentationPart.Presentation!.SlideIdList;
            var maxSlideId = slideIdList?.Elements<SlideId>().Max(s => s.Id?.Value) ?? 255U;

            var slideId = new SlideId
            {
                Id = maxSlideId + 1,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            };
            slideIdList?.Append(slideId);

            presentation.Save();
            var slideCount = slideIdList?.Elements<SlideId>().Count() ?? 1;
            return new DocumentResult(true, $"Slide {slideCount} added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add slide: {ex.Message}");
        }
    }

    public DocumentResult AddTitle(string filePath, int slideIndex, string title, string? subtitle = null, TextFormatting? textFormat = null)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            // Add title shape
            var titleShape = CreateTextShape(
                title,
                EmusPerInch / 2,
                EmusPerInch / 2,
                DefaultSlideWidth - EmusPerInch,
                EmusPerInch,
                textFormat ?? new TextFormatting(Bold: true, FontSize: 44),
                "center"
            );
            shapeTree.Append(titleShape);

            // Add subtitle if provided
            if (!string.IsNullOrEmpty(subtitle))
            {
                var subtitleShape = CreateTextShape(
                    subtitle,
                    EmusPerInch / 2,
                    EmusPerInch * 2,
                    DefaultSlideWidth - EmusPerInch,
                    EmusPerInch / 2,
                    textFormat ?? new TextFormatting(FontSize: 24),
                    "center"
                );
                shapeTree.Append(subtitleShape);
            }

            slidePart.Slide!.Save();
            return new DocumentResult(true, "Title added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add title: {ex.Message}");
        }
    }

    public DocumentResult AddTextBox(string filePath, int slideIndex, string text, TextBoxOptions options)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            var textShape = CreateTextShape(
                text,
                options.X,
                options.Y,
                options.Width,
                options.Height,
                options.TextFormat,
                options.Alignment,
                options.BackgroundColor,
                options.BorderColor,
                options.BorderWidth,
                options.VerticalAlignment,
                options.MarginLeftInches,
                options.MarginRightInches,
                options.MarginTopInches,
                options.MarginBottomInches,
                options.WordWrap,
                options.AutoFit,
                options.Rotation,
                options.GradientFill
            );
            shapeTree.Append(textShape);

            slidePart.Slide!.Save();
            return new DocumentResult(true, "Text box added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add text box: {ex.Message}");
        }
    }

    public DocumentResult AddRichTextBox(string filePath, int slideIndex, TextBoxOptions options)
    {
        try
        {
            if (options.Paragraphs == null || options.Paragraphs.Length == 0)
                return new DocumentResult(false, "Paragraphs are required for rich text box");

            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            var richTextShape = CreateRichTextShape(options);
            shapeTree.Append(richTextShape);

            slidePart.Slide!.Save();
            return new DocumentResult(true, "Rich text box added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add rich text box: {ex.Message}");
        }
    }

    public DocumentResult AddBulletPoints(string filePath, int slideIndex, string[] points, TextBoxOptions options)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            var bulletShape = CreateBulletShape(points, options);
            shapeTree.Append(bulletShape);

            slidePart.Slide!.Save();
            return new DocumentResult(true, $"Bullet list with {points.Length} items added", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add bullet points: {ex.Message}");
        }
    }

    public DocumentResult AddImage(string filePath, int slideIndex, string imagePath, long x, long y, ImageOptions? options = null)
    {
        try
        {
            if (!File.Exists(imagePath))
                return new DocumentResult(false, $"Image file not found: {imagePath}");

            options ??= new ImageOptions();

            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var imagePart = slidePart.AddImagePart(GetImagePartType(imagePath));
            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }

            var relationshipId = slidePart.GetIdOfPart(imagePart);
            var picture = CreatePicture(relationshipId, x, y, options);

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            shapeTree?.Append(picture);

            slidePart.Slide!.Save();
            return new DocumentResult(true, "Image added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add image: {ex.Message}");
        }
    }

    public DocumentResult AddImageFromBase64(string filePath, int slideIndex, string base64Data, string mimeType, long x, long y, ImageOptions? options = null)
    {
        try
        {
            options ??= new ImageOptions();

            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);

            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var partType = mimeType.ToLowerInvariant() switch
            {
                "image/png" or "png" => ImagePartType.Png,
                "image/jpeg" or "image/jpg" or "jpeg" or "jpg" => ImagePartType.Jpeg,
                "image/gif" or "gif" => ImagePartType.Gif,
                "image/bmp" or "bmp" => ImagePartType.Bmp,
                "image/svg+xml" or "svg" => ImagePartType.Svg,
                _ => ImagePartType.Png
            };

            var imagePart = slidePart.AddImagePart(partType);
            var imageBytes = Convert.FromBase64String(base64Data);
            using (var ms = new MemoryStream(imageBytes))
            {
                imagePart.FeedData(ms);
            }

            var relationshipId = slidePart.GetIdOfPart(imagePart);
            var picture = CreatePicture(relationshipId, x, y, options);

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            shapeTree?.Append(picture);

            slidePart.Slide!.Save();
            return new DocumentResult(true, "Base64 image added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add base64 image: {ex.Message}");
        }
    }

    public DocumentResult SetShapeZOrder(string filePath, int slideIndex, int shapeIndex, string position)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);

            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            // Get all child shapes (skip the first two: NonVisualGroupShapeProperties + GroupShapeProperties)
            var shapes = shapeTree.ChildElements
                .Where(e => e is P.Shape || e is P.Picture || e is P.GroupShape || e is P.ConnectionShape || e is P.GraphicFrame)
                .ToList();

            if (shapeIndex < 0 || shapeIndex >= shapes.Count)
                return new DocumentResult(false, $"Shape index {shapeIndex} out of range (0-{shapes.Count - 1})");

            var shape = shapes[shapeIndex];
            shape.Remove();

            switch (position.ToLowerInvariant())
            {
                case "front" or "top":
                    shapeTree.Append(shape);
                    break;
                case "back" or "bottom":
                    var firstShape = shapeTree.ChildElements
                        .FirstOrDefault(e => e is P.Shape || e is P.Picture || e is P.GroupShape || e is P.ConnectionShape || e is P.GraphicFrame);
                    if (firstShape != null)
                        shapeTree.InsertBefore(shape, firstShape);
                    else
                        shapeTree.Append(shape);
                    break;
                case "forward":
                    if (shapeIndex < shapes.Count - 1)
                    {
                        var nextShape = shapes[shapeIndex + 1];
                        shapeTree.InsertAfter(shape, nextShape);
                    }
                    else
                    {
                        shapeTree.Append(shape);
                    }
                    break;
                case "backward":
                    if (shapeIndex > 0)
                    {
                        var prevShape = shapes[shapeIndex - 1];
                        shapeTree.InsertBefore(shape, prevShape);
                    }
                    else
                    {
                        var first = shapeTree.ChildElements
                            .FirstOrDefault(e => e is P.Shape || e is P.Picture || e is P.GroupShape || e is P.ConnectionShape || e is P.GraphicFrame);
                        if (first != null)
                            shapeTree.InsertBefore(shape, first);
                        else
                            shapeTree.Append(shape);
                    }
                    break;
                default:
                    shapeTree.Append(shape);
                    return new DocumentResult(false, $"Unknown position: {position}", Suggestion: "Use: front, back, forward, backward");
            }

            slidePart.Slide!.Save();
            return new DocumentResult(true, $"Shape {shapeIndex} moved to {position}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to set z-order: {ex.Message}");
        }
    }

    public DocumentResult ReorderShape(string filePath, int slideIndex, int fromIndex, int toIndex)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);

            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            var shapes = shapeTree.ChildElements
                .Where(e => e is P.Shape || e is P.Picture || e is P.GroupShape || e is P.ConnectionShape || e is P.GraphicFrame)
                .ToList();

            if (fromIndex < 0 || fromIndex >= shapes.Count)
                return new DocumentResult(false, $"From index {fromIndex} out of range (0-{shapes.Count - 1})");
            if (toIndex < 0 || toIndex >= shapes.Count)
                return new DocumentResult(false, $"To index {toIndex} out of range (0-{shapes.Count - 1})");

            var shape = shapes[fromIndex];
            shape.Remove();

            // Re-fetch after removal
            var remainingShapes = shapeTree.ChildElements
                .Where(e => e is P.Shape || e is P.Picture || e is P.GroupShape || e is P.ConnectionShape || e is P.GraphicFrame)
                .ToList();

            if (toIndex >= remainingShapes.Count)
            {
                shapeTree.Append(shape);
            }
            else
            {
                shapeTree.InsertBefore(shape, remainingShapes[toIndex]);
            }

            slidePart.Slide!.Save();
            return new DocumentResult(true, $"Shape moved from position {fromIndex} to {toIndex}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to reorder shape: {ex.Message}");
        }
    }

    public DocumentResult AddShape(string filePath, int slideIndex, ShapeOptions options)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            var shape = CreateShape(options);
            shapeTree.Append(shape);

            slidePart.Slide!.Save();
            return new DocumentResult(true, $"{options.ShapeType} shape added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add shape: {ex.Message}");
        }
    }

    public DocumentResult AddLine(string filePath, int slideIndex, LineOptions options)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);

            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            var lineShape = CreateLineShape(options);
            shapeTree.Append(lineShape);

            slidePart.Slide!.Save();
            return new DocumentResult(true, "Line added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add line: {ex.Message}");
        }
    }

    public DocumentResult AddConnector(string filePath, int slideIndex, ConnectorOptions options)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);

            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            var connector = CreateConnectorShape(options);
            shapeTree.Append(connector);

            slidePart.Slide!.Save();
            return new DocumentResult(true, $"{options.ConnectorType} connector added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add connector: {ex.Message}");
        }
    }

    public DocumentResult AddGroupShape(string filePath, int slideIndex, long x, long y, long width, long height, GroupShapeItem[] items)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);

            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            var groupShape = CreateGroupShape(x, y, width, height, items, slidePart);
            shapeTree.Append(groupShape);

            slidePart.Slide!.Save();
            return new DocumentResult(true, $"Group shape with {items.Length} items added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add group shape: {ex.Message}");
        }
    }

    public DocumentResult SetSlideBackgroundGradient(string filePath, int slideIndex, GradientFillOptions gradient)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);

            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var commonSlideData = slidePart.Slide!.CommonSlideData;
            if (commonSlideData == null)
                return new DocumentResult(false, "Common slide data not found");

            var gradientFill = CreateGradientFill(gradient);
            var background = new P.Background(
                new P.BackgroundProperties(gradientFill)
            );

            var existingBackground = commonSlideData.Elements<P.Background>().FirstOrDefault();
            existingBackground?.Remove();

            commonSlideData.InsertAt(background, 0);

            slidePart.Slide!.Save();
            return new DocumentResult(true, $"Slide {slideIndex} background set to gradient", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to set gradient background: {ex.Message}");
        }
    }

    public DocumentResult SetSlideSize(string filePath, string size)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var presentationPart = presentation.PresentationPart;

            if (presentationPart == null)
                return new DocumentResult(false, "Presentation part not found");

            var (cx, cy) = size.ToLowerInvariant() switch
            {
                "widescreen" or "16:9" => (12192000, 6858000),
                "standard" or "4:3" => (9144000, 6858000),
                "widescreen16x10" or "16:10" => (10972800, 6858000),
                "a4" => (10691813, 7559675),
                "letter" => (10058400, 7772400),
                _ => (12192000, 6858000)
            };

            var slideSize = presentationPart.Presentation!.SlideSize;
            if (slideSize != null)
            {
                slideSize.Cx = cx;
                slideSize.Cy = cy;
            }

            presentation.Save();
            return new DocumentResult(true, $"Slide size set to {size}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to set slide size: {ex.Message}");
        }
    }

    public DocumentResult AddTable(string filePath, int slideIndex, string[][] data, long x, long y, long width, long height)
    {
        try
        {
            if (data.Length == 0)
                return new DocumentResult(false, "Table data cannot be empty");

            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                return new DocumentResult(false, "Shape tree not found");

            var graphicFrame = CreateTableGraphicFrame(data, x, y, width, height);
            shapeTree.Append(graphicFrame);

            slidePart.Slide!.Save();
            return new DocumentResult(true, $"Table with {data.Length} rows added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add table: {ex.Message}");
        }
    }

    public DocumentResult SetSlideBackground(string filePath, int slideIndex, string color)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var commonSlideData = slidePart.Slide!.CommonSlideData;
            if (commonSlideData == null)
                return new DocumentResult(false, "Common slide data not found");

            var background = new P.Background(
                new P.BackgroundProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex { Val = color.TrimStart('#') }
                    )
                )
            );

            var existingBackground = commonSlideData.Elements<P.Background>().FirstOrDefault();
            if (existingBackground != null)
            {
                existingBackground.Remove();
            }

            commonSlideData.InsertAt(background, 0);

            slidePart.Slide!.Save();
            return new DocumentResult(true, $"Slide {slideIndex} background set to #{color.TrimStart('#')}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to set background: {ex.Message}");
        }
    }

    public DocumentResult DeleteSlide(string filePath, int slideIndex)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var presentationPart = presentation.PresentationPart;
            
            if (presentationPart == null)
                return new DocumentResult(false, "Presentation part not found");

            var slideIdList = presentationPart.Presentation!.SlideIdList;
            var slideIds = slideIdList?.Elements<SlideId>().ToList() ?? [];
            
            if (slideIndex < 0 || slideIndex >= slideIds.Count)
                return new DocumentResult(false, $"Slide index {slideIndex} is out of range");

            var slideId = slideIds[slideIndex];
            var relationshipId = slideId.RelationshipId?.Value;

            if (relationshipId != null)
            {
                var slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
                presentationPart.DeletePart(slidePart);
            }

            slideId.Remove();
            presentation.Save();
            return new DocumentResult(true, $"Slide {slideIndex} deleted", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to delete slide: {ex.Message}");
        }
    }

    public DocumentResult DuplicateSlide(string filePath, int sourceIndex)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var presentationPart = presentation.PresentationPart;
            
            if (presentationPart == null)
                return new DocumentResult(false, "Presentation part not found");

            var sourceSlidePart = GetSlidePart(presentation, sourceIndex);
            if (sourceSlidePart == null)
                return new DocumentResult(false, $"Source slide {sourceIndex} not found");

            var newSlidePart = presentationPart.AddNewPart<SlidePart>();
            
            using (var sourceStream = sourceSlidePart.GetStream())
            {
                newSlidePart.FeedData(sourceStream);
            }

            // Copy layout relationship
            foreach (var part in sourceSlidePart.Parts)
            {
                if (part.OpenXmlPart is SlideLayoutPart)
                {
                    newSlidePart.AddPart(part.OpenXmlPart, part.RelationshipId);
                }
            }

            var slideIdList = presentationPart.Presentation!.SlideIdList;
            var maxSlideId = slideIdList?.Elements<SlideId>().Max(s => s.Id?.Value) ?? 255U;

            slideIdList?.Append(new SlideId
            {
                Id = maxSlideId + 1,
                RelationshipId = presentationPart.GetIdOfPart(newSlidePart)
            });

            presentation.Save();
            return new DocumentResult(true, $"Slide {sourceIndex} duplicated", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to duplicate slide: {ex.Message}");
        }
    }

    public DocumentResult ReorderSlide(string filePath, int fromIndex, int toIndex)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var presentationPart = presentation.PresentationPart;
            
            if (presentationPart == null)
                return new DocumentResult(false, "Presentation part not found");

            var slideIdList = presentationPart.Presentation!.SlideIdList;
            var slideIds = slideIdList?.Elements<SlideId>().ToList() ?? [];
            
            if (fromIndex < 0 || fromIndex >= slideIds.Count)
                return new DocumentResult(false, $"From index {fromIndex} is out of range");
            
            if (toIndex < 0 || toIndex >= slideIds.Count)
                return new DocumentResult(false, $"To index {toIndex} is out of range");

            var slideId = slideIds[fromIndex];
            slideId.Remove();

            var referenceSlide = slideIds.ElementAtOrDefault(toIndex);
            if (referenceSlide != null && toIndex < fromIndex)
            {
                slideIdList?.InsertBefore(slideId, referenceSlide);
            }
            else if (referenceSlide != null)
            {
                slideIdList?.InsertAfter(slideId, referenceSlide);
            }
            else
            {
                slideIdList?.Append(slideId);
            }

            presentation.Save();
            return new DocumentResult(true, $"Slide moved from position {fromIndex} to {toIndex}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to reorder slide: {ex.Message}");
        }
    }

    public ContentResult GetSlideText(string filePath, int slideIndex)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var presentation = PresentationDocument.Open(filePath, false);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new ContentResult(false, null, $"Slide {slideIndex} not found");

            var textContent = ExtractSlideText(slidePart);
            return new ContentResult(true, textContent);
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get slide text: {ex.Message}");
        }
    }

    public ContentResult GetAllSlidesText(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var presentation = PresentationDocument.Open(filePath, false);
            var presentationPart = presentation.PresentationPart;
            
            if (presentationPart == null)
                return new ContentResult(false, null, "Presentation part not found");

            var slideIds = presentationPart.Presentation!.SlideIdList?.Elements<SlideId>().ToList() ?? [];
            var sb = new StringBuilder();

            for (int i = 0; i < slideIds.Count; i++)
            {
                sb.AppendLine($"=== Slide {i + 1} ===");
                var slidePart = GetSlidePart(presentation, i);
                if (slidePart != null)
                {
                    sb.AppendLine(ExtractSlideText(slidePart));
                }
                sb.AppendLine();
            }

            return new ContentResult(true, sb.ToString().TrimEnd(), TotalPages: slideIds.Count);
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get presentation text: {ex.Message}");
        }
    }

    public ContentResult GetSlideCount(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var presentation = PresentationDocument.Open(filePath, false);
            var slideCount = presentation.PresentationPart?.Presentation!.SlideIdList?.Elements<SlideId>().Count() ?? 0;
            
            return new ContentResult(true, slideCount.ToString(), TotalPages: slideCount);
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get slide count: {ex.Message}");
        }
    }

    public DocumentResult AddSpeakerNotes(string filePath, int slideIndex, string notes)
    {
        try
        {
            using var presentation = PresentationDocument.Open(filePath, true);
            var slidePart = GetSlidePart(presentation, slideIndex);
            
            if (slidePart == null)
                return new DocumentResult(false, $"Slide {slideIndex} not found");

            var notesSlidePart = slidePart.NotesSlidePart;
            if (notesSlidePart == null)
            {
                notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
                notesSlidePart.NotesSlide = CreateNotesSlide();
            }

            var textBody = notesSlidePart.NotesSlide!.Descendants<P.TextBody>().FirstOrDefault();
            if (textBody != null)
            {
                textBody.RemoveAllChildren<A.Paragraph>();
                textBody.Append(new A.Paragraph(
                    new A.Run(new A.Text(notes))
                ));
            }

            notesSlidePart.NotesSlide!.Save();
            return new DocumentResult(true, $"Speaker notes added to slide {slideIndex}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add speaker notes: {ex.Message}");
        }
    }

    public IList<ImageExtractionResult> ExtractImages(string filePath)
    {
        var results = new List<ImageExtractionResult>();
        try
        {
            if (!File.Exists(filePath)) return results;

            using var presentation = PresentationDocument.Open(filePath, false);
            var presentationPart = presentation.PresentationPart;
            if (presentationPart == null) return results;

            var slideIds = presentationPart.Presentation!.SlideIdList?.Elements<SlideId>().ToList() ?? [];
            int imageIndex = 0;

            for (int slideIdx = 0; slideIdx < slideIds.Count; slideIdx++)
            {
                var slideId = slideIds[slideIdx];
                if (slideId.RelationshipId?.Value == null) continue;

                var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId.Value);
                var shapeTree = slidePart.Slide!.CommonSlideData?.ShapeTree;
                if (shapeTree == null) continue;

                // Slide text provides context for AI captioning
                var slideText = ExtractSlideText(slidePart);

                foreach (var picture in shapeTree.Descendants<P.Picture>())
                {
                    try
                    {
                        var blip = picture.Descendants<A.Blip>().FirstOrDefault();
                        if (blip?.Embed == null) continue;

                        var imagePart = (ImagePart)slidePart.GetPartById(blip.Embed!);

                        // Alt text from NonVisualDrawingProperties
                        var nvProps = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
                        var altText = nvProps?.Description?.Value ?? nvProps?.Name?.Value ?? string.Empty;

                        // Image bytes as base64
                        using var imgStream = imagePart.GetStream();
                        using var ms = new MemoryStream();
                        imgStream.CopyTo(ms);
                        var imageBase64 = Convert.ToBase64String(ms.ToArray());

                        // Dimensions from ShapeProperties.Transform2D.Extents
                        var extents = picture.ShapeProperties?.Transform2D?.Extents;
                        int? widthPx = extents?.Cx != null ? (int)(extents.Cx.Value * 96.0 / EmusPerInch) : null;
                        int? heightPx = extents?.Cy != null ? (int)(extents.Cy.Value * 96.0 / EmusPerInch) : null;

                        results.Add(new ImageExtractionResult(
                            Index: imageIndex++,
                            MimeType: imagePart.ContentType,
                            ImageBase64: imageBase64,
                            AltText: altText,
                            ContextBefore: slideText,
                            ContextAfter: string.Empty,
                            WidthPx: widthPx,
                            HeightPx: heightPx,
                            PageOrSlideNumber: slideIdx + 1
                        ));
                    }
                    catch
                    {
                        // Skip problematic images
                    }
                }
            }
        }
        catch
        {
            // Return whatever was collected before the error
        }
        return results;
    }

    #region Private Helper Methods

    private static SlidePart? GetSlidePart(PresentationDocument presentation, int slideIndex)
    {
        var presentationPart = presentation.PresentationPart;
        var slideIds = presentationPart?.Presentation!.SlideIdList?.Elements<SlideId>().ToList() ?? [];
        
        if (slideIndex < 0 || slideIndex >= slideIds.Count)
            return null;

        var slideId = slideIds[slideIndex];
        return slideId.RelationshipId?.Value != null
            ? (SlidePart)presentationPart!.GetPartById(slideId.RelationshipId.Value)
            : null;
    }

    private static P.Slide CreateSlide(SlideLayoutOptions? options = null)
    {
        var slide = new P.Slide(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()
                    ),
                    new P.GroupShapeProperties(
                        new A.TransformGroup(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 0L, Cy = 0L },
                            new A.ChildOffset { X = 0L, Y = 0L },
                            new A.ChildExtents { Cx = 0L, Cy = 0L }
                        )
                    )
                )
            ),
            new P.ColorMapOverride(new A.MasterColorMapping())
        );

        if (!string.IsNullOrEmpty(options?.BackgroundColor))
        {
            var background = new P.Background(
                new P.BackgroundProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex { Val = options.BackgroundColor.TrimStart('#') }
                    )
                )
            );
            slide.CommonSlideData!.InsertAt(background, 0);
        }
        else if (options?.GradientBackground != null)
        {
            var gradientFill = CreateGradientFill(options.GradientBackground);
            var background = new P.Background(
                new P.BackgroundProperties(gradientFill)
            );
            slide.CommonSlideData!.InsertAt(background, 0);
        }

        return slide;
    }

    private static P.SlideMaster CreateSlideMaster()
    {
        return new P.SlideMaster(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()
                    ),
                    new P.GroupShapeProperties(
                        new A.TransformGroup(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 0L, Cy = 0L },
                            new A.ChildOffset { X = 0L, Y = 0L },
                            new A.ChildExtents { Cx = 0L, Cy = 0L }
                        )
                    )
                )
            ),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            }
        );
    }

    private static P.SlideLayout CreateSlideLayout()
    {
        return new P.SlideLayout(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()
                    ),
                    new P.GroupShapeProperties(
                        new A.TransformGroup(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 0L, Cy = 0L },
                            new A.ChildOffset { X = 0L, Y = 0L },
                            new A.ChildExtents { Cx = 0L, Cy = 0L }
                        )
                    )
                )
            ),
            new P.ColorMapOverride(new A.MasterColorMapping())
        )
        { Type = SlideLayoutValues.Blank };
    }

    private static A.Theme CreateDefaultTheme()
    {
        return new A.Theme(
            new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                    new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = "44546A" }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = "E7E6E6" }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = "4472C4" }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "ED7D31" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "A5A5A5" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "FFC000" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "70AD47" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" })
                )
                { Name = "Office Theme" },
                new A.FontScheme(
                    new A.MajorFont(
                        new A.LatinFont { Typeface = "Calibri Light" },
                        new A.EastAsianFont { Typeface = string.Empty },
                        new A.ComplexScriptFont { Typeface = string.Empty }
                    ),
                    new A.MinorFont(
                        new A.LatinFont { Typeface = "Calibri" },
                        new A.EastAsianFont { Typeface = string.Empty },
                        new A.ComplexScriptFont { Typeface = string.Empty }
                    )
                )
                { Name = "Office" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }
                            ),
                            new A.LinearGradientFill { Angle = 5400000, Scaled = false }
                        ),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }
                            ),
                            new A.LinearGradientFill { Angle = 5400000, Scaled = false }
                        )
                    ),
                    new A.LineStyleList(
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 9525 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 25400 },
                        new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = 38100 }
                    ),
                    new A.EffectStyleList(
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList())
                    ),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                        new A.GradientFill(
                            new A.GradientStopList(
                                new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                                new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }
                            ),
                            new A.LinearGradientFill { Angle = 5400000, Scaled = false }
                        )
                    )
                )
                { Name = "Office" }
            ),
            new A.ObjectDefaults(),
            new A.ExtraColorSchemeList()
        )
        { Name = "Office Theme" };
    }

    private static P.Shape CreateTextShape(string text, long x, long y, long width, long height,
        TextFormatting? format = null, string alignment = "left", string? backgroundColor = null, string? borderColor = null,
        double borderWidth = 1.0, string verticalAlignment = "Top",
        double marginLeft = 0.1, double marginRight = 0.1, double marginTop = 0.05, double marginBottom = 0.05,
        bool wordWrap = true, string autoFit = "None", double rotation = 0.0,
        GradientFillOptions? gradientFill = null)
    {
        var shapeId = NextShapeId();

        var bodyProperties = new A.BodyProperties
        {
            Wrap = wordWrap ? A.TextWrappingValues.Square : A.TextWrappingValues.None,
            LeftInset = (int)(marginLeft * EmusPerInch),
            RightInset = (int)(marginRight * EmusPerInch),
            TopInset = (int)(marginTop * EmusPerInch),
            BottomInset = (int)(marginBottom * EmusPerInch),
            Anchor = GetVerticalAnchor(verticalAlignment)
        };

        if (autoFit.Equals("ShrinkText", StringComparison.OrdinalIgnoreCase))
            bodyProperties.Append(new A.NormalAutoFit());
        else if (autoFit.Equals("ResizeShape", StringComparison.OrdinalIgnoreCase))
            bodyProperties.Append(new A.ShapeAutoFit());

        var transform = new A.Transform2D(
            new A.Offset { X = x, Y = y },
            new A.Extents { Cx = width, Cy = height }
        );
        if (rotation != 0.0)
            transform.Rotation = (int)(rotation * 60000);

        var shapeProperties = new P.ShapeProperties(
            transform,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
        );

        // Fill
        if (gradientFill != null)
            shapeProperties.Append(CreateGradientFill(gradientFill));
        else if (!string.IsNullOrEmpty(backgroundColor))
            shapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = backgroundColor.TrimStart('#') }));
        else
            shapeProperties.Append(new A.NoFill());

        // Border
        if (!string.IsNullOrEmpty(borderColor))
        {
            shapeProperties.Append(new A.Outline(
                new A.SolidFill(new A.RgbColorModelHex { Val = borderColor.TrimStart('#') })
            )
            { Width = (int)(borderWidth * 12700) });
        }

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            shapeProperties,
            new P.TextBody(
                bodyProperties,
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties { Alignment = GetTextAlignment(alignment) },
                    CreateTextRun(text, format)
                )
            )
        );

        return shape;
    }

    private static A.Run CreateTextRun(string text, TextFormatting? format)
    {
        var run = new A.Run(new A.Text(text));
        
        if (format != null)
        {
            var runProps = new A.RunProperties { Language = "en-US" };
            
            if (format.Bold)
                runProps.Bold = true;
            
            if (format.Italic)
                runProps.Italic = true;
            
            if (format.Underline)
                runProps.Underline = A.TextUnderlineValues.Single;
            
            if (format.Strikethrough)
                runProps.Strike = A.TextStrikeValues.SingleStrike;
            
            if (format.FontSize.HasValue)
                runProps.FontSize = format.FontSize.Value * 100;
            
            if (!string.IsNullOrEmpty(format.FontColor))
            {
                runProps.Append(new A.SolidFill(
                    new A.RgbColorModelHex { Val = format.FontColor.TrimStart('#') }
                ));
            }
            
            if (!string.IsNullOrEmpty(format.FontName))
            {
                runProps.Append(new A.LatinFont { Typeface = format.FontName });
            }

            run.InsertAt(runProps, 0);
        }

        return run;
    }

    private static A.TextAlignmentTypeValues GetTextAlignment(string alignment) => alignment.ToLowerInvariant() switch
    {
        "center" => A.TextAlignmentTypeValues.Center,
        "right" => A.TextAlignmentTypeValues.Right,
        "justify" => A.TextAlignmentTypeValues.Justified,
        _ => A.TextAlignmentTypeValues.Left
    };

    private static P.Shape CreateBulletShape(string[] points, TextBoxOptions options)
    {
        var shapeId = NextShapeId();
        
        var textBody = new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle()
        );

        foreach (var point in points)
        {
            var paragraph = new A.Paragraph(
                new A.ParagraphProperties(
                    new A.BulletFont { Typeface = "Arial" },
                    new A.CharacterBullet { Char = "\u2022" }
                )
                { LeftMargin = 228600, Indent = -228600 },
                CreateTextRun(point, options.TextFormat)
            );
            textBody.Append(paragraph);
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"BulletList {shapeId}" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = options.X, Y = options.Y },
                    new A.Extents { Cx = options.Width, Cy = options.Height }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill()
            ),
            textBody
        );
    }

    private static P.Picture CreatePicture(string relationshipId, long x, long y, ImageOptions options)
    {
        var pictureId = NextShapeId();

        // Determine crop shape (default Rectangle, use Ellipse for circular avatars)
        var cropShapeType = MapShapeType(options.CropShape ?? "Rectangle");

        var transform = new A.Transform2D(
            new A.Offset { X = x, Y = y },
            new A.Extents { Cx = options.WidthEmu, Cy = options.HeightEmu }
        );
        if (options.Rotation != 0.0)
            transform.Rotation = (int)(options.Rotation * 60000);

        var shapeProperties = new P.ShapeProperties(
            transform,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = cropShapeType }
        );

        // Border/outline
        if (!string.IsNullOrEmpty(options.BorderColor) && options.BorderWidth > 0)
        {
            shapeProperties.Append(new A.Outline(
                new A.SolidFill(new A.RgbColorModelHex { Val = options.BorderColor.TrimStart('#') })
            )
            { Width = (int)(options.BorderWidth * 12700) });
        }

        // Shadow
        if (options.HasShadow)
        {
            shapeProperties.Append(CreateShadowEffect());
        }

        // 3D perspective
        if (options.Perspective3DAngleX != 0.0 || options.Perspective3DAngleY != 0.0)
        {
            shapeProperties.Append(Create3DScene(options.Perspective3DAngleX, options.Perspective3DAngleY));
        }

        return new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = pictureId, Name = $"Picture {pictureId}", Description = options.AltText ?? string.Empty },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())
            ),
            shapeProperties
        );
    }

    private static P.Shape CreateShape(ShapeOptions options)
    {
        var shapeId = NextShapeId();

        var shapeType = MapShapeType(options.ShapeType);

        var transform = new A.Transform2D(
            new A.Offset { X = options.X, Y = options.Y },
            new A.Extents { Cx = options.Width, Cy = options.Height }
        );
        if (options.Rotation != 0.0)
            transform.Rotation = (int)(options.Rotation * 60000);

        var presetGeometry = new A.PresetGeometry(new A.AdjustValueList()) { Preset = shapeType };

        // Corner radius for rounded rectangles
        if (options.CornerRadiusPt.HasValue && shapeType == A.ShapeTypeValues.RoundRectangle)
        {
            var adjustList = new A.AdjustValueList(
                new A.ShapeGuide { Name = "adj", Formula = $"val {options.CornerRadiusPt.Value * 12700}" }
            );
            presetGeometry = new A.PresetGeometry(adjustList) { Preset = shapeType };
        }

        var shapeProperties = new P.ShapeProperties(transform, presetGeometry);

        // Fill
        if (options.NoFill)
        {
            shapeProperties.Append(new A.NoFill());
        }
        else if (options.GradientFill != null)
        {
            shapeProperties.Append(CreateGradientFill(options.GradientFill));
        }
        else if (!string.IsNullOrEmpty(options.FillColor))
        {
            var fillColor = new A.RgbColorModelHex { Val = options.FillColor.TrimStart('#') };
            if (options.TransparencyPercent > 0)
                fillColor.Append(new A.Alpha { Val = (int)((100 - options.TransparencyPercent) * 1000) });
            shapeProperties.Append(new A.SolidFill(fillColor));
        }

        // Border/Outline
        if (!string.IsNullOrEmpty(options.BorderColor) || options.BorderWidth > 0)
        {
            var outline = new A.Outline { Width = (int)(options.BorderWidth * 12700) };
            if (!string.IsNullOrEmpty(options.BorderColor))
                outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = options.BorderColor.TrimStart('#') }));
            
            ApplyDashStyle(outline, options.DashStyle);
            shapeProperties.Append(outline);
        }

        // Shadow
        if (options.HasShadow)
        {
            shapeProperties.Append(CreateShadowEffect());
        }

        // 3D perspective
        if (options.Perspective3DAngleX != 0.0 || options.Perspective3DAngleY != 0.0)
        {
            shapeProperties.Append(Create3DScene(options.Perspective3DAngleX, options.Perspective3DAngleY));
        }

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Shape {shapeId}" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            shapeProperties
        );

        // Text inside shape
        if (options.Paragraphs != null && options.Paragraphs.Length > 0)
        {
            var textBody = CreateRichTextBody(options.Paragraphs, options.TextAlignment, options.VerticalTextAlignment,
                options.MarginLeftInches, options.MarginRightInches, options.MarginTopInches, options.MarginBottomInches);
            shape.Append(textBody);
        }
        else if (!string.IsNullOrEmpty(options.Text))
        {
            var bodyProps = new A.BodyProperties
            {
                Anchor = GetVerticalAnchor(options.VerticalTextAlignment),
                LeftInset = (int)(options.MarginLeftInches * EmusPerInch),
                RightInset = (int)(options.MarginRightInches * EmusPerInch),
                TopInset = (int)(options.MarginTopInches * EmusPerInch),
                BottomInset = (int)(options.MarginBottomInches * EmusPerInch)
            };

            shape.Append(new P.TextBody(
                bodyProps,
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties { Alignment = GetTextAlignment(options.TextAlignment) },
                    CreateTextRun(options.Text, options.TextFormat)
                )
            ));
        }

        return shape;
    }

    private static P.ConnectionShape CreateConnectorShape(ConnectorOptions options)
    {
        var shapeId = NextShapeId();

        // Compute position and extents from endpoints
        var x = Math.Min(options.X1, options.X2);
        var y = Math.Min(options.Y1, options.Y2);
        var cx = Math.Abs(options.X2 - options.X1);
        var cy = Math.Abs(options.Y2 - options.Y1);
        if (cx == 0) cx = 1;
        if (cy == 0) cy = 1;

        var flipH = options.X2 < options.X1;
        var flipV = options.Y2 < options.Y1;

        var connectorType = options.ConnectorType.ToLowerInvariant() switch
        {
            "elbow" or "bent" => A.ShapeTypeValues.BentConnector3,
            "curved" => A.ShapeTypeValues.CurvedConnector3,
            _ => A.ShapeTypeValues.StraightConnector1
        };

        var transform = new A.Transform2D(
            new A.Offset { X = x, Y = y },
            new A.Extents { Cx = cx, Cy = cy }
        );
        if (flipH) transform.HorizontalFlip = true;
        if (flipV) transform.VerticalFlip = true;

        var shapeProperties = new P.ShapeProperties(
            transform,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = connectorType }
        );

        // Line styling
        var outline = new A.Outline { Width = (int)(options.LineWidth * 12700) };
        if (!string.IsNullOrEmpty(options.LineColor))
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = options.LineColor.TrimStart('#') }));
        else
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }));

        ApplyDashStyle(outline, options.DashStyle);
        ApplyArrowHeads(outline, options.StartArrow, options.EndArrow);
        shapeProperties.Append(outline);

        return new P.ConnectionShape(
            new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Connector {shapeId}" },
                new P.NonVisualConnectorShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            shapeProperties
        );
    }

    private static P.Shape CreateLineShape(LineOptions options)
    {
        var shapeId = NextShapeId();

        var x = Math.Min(options.X1, options.X2);
        var y = Math.Min(options.Y1, options.Y2);
        var cx = Math.Abs(options.X2 - options.X1);
        var cy = Math.Abs(options.Y2 - options.Y1);
        if (cx == 0) cx = 1;
        if (cy == 0) cy = 1;

        var flipH = options.X2 < options.X1;
        var flipV = options.Y2 < options.Y1;

        var transform = new A.Transform2D(
            new A.Offset { X = x, Y = y },
            new A.Extents { Cx = cx, Cy = cy }
        );
        if (flipH) transform.HorizontalFlip = true;
        if (flipV) transform.VerticalFlip = true;

        var shapeProperties = new P.ShapeProperties(
            transform,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Line }
        );

        shapeProperties.Append(new A.NoFill());

        var outline = new A.Outline { Width = (int)(options.LineWidth * 12700) };
        if (!string.IsNullOrEmpty(options.LineColor))
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = options.LineColor.TrimStart('#') }));
        else
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }));

        ApplyDashStyle(outline, options.DashStyle);
        ApplyArrowHeads(outline, options.StartArrow, options.EndArrow);
        shapeProperties.Append(outline);

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Line {shapeId}" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            shapeProperties
        );
    }

    private static P.GroupShape CreateGroupShape(long x, long y, long width, long height, GroupShapeItem[] items, SlidePart slidePart)
    {
        var groupId = NextShapeId();

        var groupShape = new P.GroupShape(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = groupId, Name = $"Group {groupId}" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.GroupShapeProperties(
                new A.TransformGroup(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = width, Cy = height },
                    new A.ChildOffset { X = x, Y = y },
                    new A.ChildExtents { Cx = width, Cy = height }
                )
            )
        );

        foreach (var item in items)
        {
            switch (item.ItemType.ToLowerInvariant())
            {
                case "shape":
                    var shapeOpts = new ShapeOptions(
                        ShapeType: item.ShapeType ?? "Rectangle",
                        X: item.X, Y: item.Y, Width: item.Width, Height: item.Height,
                        FillColor: item.FillColor, BorderColor: item.BorderColor,
                        BorderWidth: item.BorderWidth, Text: item.Text, TextFormat: item.TextFormat
                    );
                    groupShape.Append(CreateShape(shapeOpts));
                    break;

                case "textbox":
                    var textShape = CreateTextShape(
                        item.Text ?? string.Empty,
                        item.X, item.Y, item.Width, item.Height,
                        item.TextFormat
                    );
                    groupShape.Append(textShape);
                    break;

                case "image" when !string.IsNullOrEmpty(item.ImagePath) && File.Exists(item.ImagePath):
                    var imagePart = slidePart.AddImagePart(GetImagePartType(item.ImagePath));
                    using (var stream = new FileStream(item.ImagePath, FileMode.Open, FileAccess.Read))
                    {
                        imagePart.FeedData(stream);
                    }
                    var relId = slidePart.GetIdOfPart(imagePart);
                    var imgOpts = new ImageOptions(WidthEmu: item.Width, HeightEmu: item.Height);
                    groupShape.Append(CreatePicture(relId, item.X, item.Y, imgOpts));
                    break;

                case "line":
                    var lineOpts = new LineOptions(
                        X1: item.X, Y1: item.Y,
                        X2: item.X + item.Width, Y2: item.Y + item.Height,
                        LineColor: item.BorderColor, LineWidth: item.BorderWidth
                    );
                    groupShape.Append(CreateLineShape(lineOpts));
                    break;
            }
        }

        return groupShape;
    }

    private static P.Shape CreateRichTextShape(TextBoxOptions options)
    {
        var shapeId = NextShapeId();

        var transform = new A.Transform2D(
            new A.Offset { X = options.X, Y = options.Y },
            new A.Extents { Cx = options.Width, Cy = options.Height }
        );
        if (options.Rotation != 0.0)
            transform.Rotation = (int)(options.Rotation * 60000);

        var shapeProperties = new P.ShapeProperties(
            transform,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
        );

        if (options.GradientFill != null)
            shapeProperties.Append(CreateGradientFill(options.GradientFill));
        else if (!string.IsNullOrEmpty(options.BackgroundColor))
            shapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = options.BackgroundColor.TrimStart('#') }));
        else
            shapeProperties.Append(new A.NoFill());

        if (!string.IsNullOrEmpty(options.BorderColor))
        {
            shapeProperties.Append(new A.Outline(
                new A.SolidFill(new A.RgbColorModelHex { Val = options.BorderColor.TrimStart('#') })
            )
            { Width = (int)(options.BorderWidth * 12700) });
        }

        var textBody = CreateRichTextBody(options.Paragraphs!, options.Alignment, options.VerticalAlignment,
            options.MarginLeftInches, options.MarginRightInches, options.MarginTopInches, options.MarginBottomInches,
            options.WordWrap, options.AutoFit);

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"RichTextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            shapeProperties,
            textBody
        );
    }

    private static P.TextBody CreateRichTextBody(RichParagraph[] paragraphs, string alignment = "Left",
        string verticalAlignment = "Top", double marginLeft = 0.1, double marginRight = 0.1,
        double marginTop = 0.05, double marginBottom = 0.05, bool wordWrap = true, string autoFit = "None")
    {
        var bodyProps = new A.BodyProperties
        {
            Wrap = wordWrap ? A.TextWrappingValues.Square : A.TextWrappingValues.None,
            LeftInset = (int)(marginLeft * EmusPerInch),
            RightInset = (int)(marginRight * EmusPerInch),
            TopInset = (int)(marginTop * EmusPerInch),
            BottomInset = (int)(marginBottom * EmusPerInch),
            Anchor = GetVerticalAnchor(verticalAlignment)
        };

        if (autoFit.Equals("ShrinkText", StringComparison.OrdinalIgnoreCase))
            bodyProps.Append(new A.NormalAutoFit());
        else if (autoFit.Equals("ResizeShape", StringComparison.OrdinalIgnoreCase))
            bodyProps.Append(new A.ShapeAutoFit());

        var textBody = new P.TextBody(bodyProps, new A.ListStyle());

        foreach (var para in paragraphs)
        {
            var paragraph = new A.Paragraph();

            var paraProps = new A.ParagraphProperties
            {
                Alignment = GetTextAlignment(para.Alignment ?? alignment)
            };

            if (para.SpacingBeforePt.HasValue)
                paraProps.Append(new A.SpaceBefore(new A.SpacingPoints { Val = (int)(para.SpacingBeforePt.Value * 100) }));
            if (para.SpacingAfterPt.HasValue)
                paraProps.Append(new A.SpaceAfter(new A.SpacingPoints { Val = (int)(para.SpacingAfterPt.Value * 100) }));
            if (para.LineSpacingPercent.HasValue)
                paraProps.Append(new A.LineSpacing(new A.SpacingPercent { Val = (int)(para.LineSpacingPercent.Value * 1000) }));

            if (para.IsBullet)
            {
                paraProps.LeftMargin = 228600 * (para.IndentLevel ?? 1);
                paraProps.Indent = -228600;
                paraProps.Append(new A.BulletFont { Typeface = "Arial" });
                paraProps.Append(new A.CharacterBullet { Char = para.BulletChar ?? "•" });
            }

            paragraph.Append(paraProps);

            // Multi-run support
            if (para.Runs != null && para.Runs.Length > 0)
            {
                foreach (var run in para.Runs)
                {
                    paragraph.Append(CreateTextRun(run.Text, run.Format));
                }
            }
            else if (!string.IsNullOrEmpty(para.Text))
            {
                paragraph.Append(CreateTextRun(para.Text, null));
            }
            else
            {
                // Empty paragraph (spacing line)
                paragraph.Append(new A.EndParagraphRunProperties { Language = "en-US" });
            }

            textBody.Append(paragraph);
        }

        return textBody;
    }

    private static A.GradientFill CreateGradientFill(GradientFillOptions options)
    {
        var stopList = new A.GradientStopList();
        foreach (var stop in options.Stops)
        {
            var color = new A.RgbColorModelHex { Val = stop.Color.TrimStart('#') };
            if (stop.TransparencyPercent.HasValue && stop.TransparencyPercent > 0)
                color.Append(new A.Alpha { Val = (int)((100 - stop.TransparencyPercent.Value) * 1000) });

            stopList.Append(new A.GradientStop(color) { Position = stop.Position * 1000 });
        }

        var gradientFill = new A.GradientFill(stopList);

        if (options.GradientType.Equals("Linear", StringComparison.OrdinalIgnoreCase))
        {
            gradientFill.Append(new A.LinearGradientFill
            {
                Angle = (int)(options.Angle * 60000),
                Scaled = false
            });
        }

        return gradientFill;
    }

    private static OpenXmlElement CreateShadowEffect()
    {
        return new A.EffectList(
            new A.OuterShadow(
                new A.RgbColorModelHex(
                    new A.Alpha { Val = 40000 }
                )
                { Val = "000000" }
            )
            {
                BlurRadius = 76200L,
                Distance = 38100L,
                Direction = 2700000,
                Alignment = A.RectangleAlignmentValues.TopLeft,
                RotateWithShape = false
            }
        );
    }

    /// <summary>
    /// Creates a 3D scene with perspective camera rotation for shapes/images.
    /// This enables the stacked card perspective effect seen in modern slide designs.
    /// OpenXML represents 3D rotation via a:scene3d with camera rotation angles.
    /// </summary>
    private static A.Scene3DType Create3DScene(double rotationX, double rotationY)
    {
        // OpenXML uses 60000ths of a degree for rotation values
        var rotXVal = (int)(rotationX * 60000);
        var rotYVal = (int)(rotationY * 60000);

        return new A.Scene3DType(
            new A.Camera(
                new A.Rotation
                {
                    Latitude = rotXVal,
                    Longitude = rotYVal,
                    Revolution = 0
                }
            )
            { Preset = A.PresetCameraValues.PerspectiveFront },
            new A.LightRig
            {
                Rig = A.LightRigValues.ThreePoints,
                Direction = A.LightRigDirectionValues.Top
            }
        );
    }

    private static void ApplyDashStyle(A.Outline outline, string dashStyle)
    {
        var preset = dashStyle.ToLowerInvariant() switch
        {
            "dash" => A.PresetLineDashValues.Dash,
            "dashdot" => A.PresetLineDashValues.DashDot,
            "dot" => A.PresetLineDashValues.Dot,
            "longdash" => A.PresetLineDashValues.LargeDash,
            "longdashdot" => A.PresetLineDashValues.LargeDashDot,
            "longdashdotdot" => A.PresetLineDashValues.LargeDashDotDot,
            "sysdash" => A.PresetLineDashValues.SystemDash,
            "sysdot" => A.PresetLineDashValues.SystemDot,
            _ => (A.PresetLineDashValues?)null
        };

        if (preset.HasValue)
            outline.Append(new A.PresetDash { Val = preset.Value });
    }

    private static void ApplyArrowHeads(A.Outline outline, string? startArrow, string? endArrow)
    {
        if (!string.IsNullOrEmpty(startArrow))
        {
            outline.Append(new A.HeadEnd
            {
                Type = MapArrowType(startArrow),
                Width = A.LineEndWidthValues.Medium,
                Length = A.LineEndLengthValues.Medium
            });
        }

        if (!string.IsNullOrEmpty(endArrow))
        {
            outline.Append(new A.TailEnd
            {
                Type = MapArrowType(endArrow),
                Width = A.LineEndWidthValues.Medium,
                Length = A.LineEndLengthValues.Medium
            });
        }
    }

    private static A.LineEndValues MapArrowType(string arrowType) => arrowType.ToLowerInvariant() switch
    {
        "triangle" or "arrow" => A.LineEndValues.Triangle,
        "stealth" => A.LineEndValues.Stealth,
        "diamond" => A.LineEndValues.Diamond,
        "oval" or "circle" => A.LineEndValues.Oval,
        "open" => A.LineEndValues.Arrow,
        _ => A.LineEndValues.Triangle
    };

    private static A.TextAnchoringTypeValues GetVerticalAnchor(string alignment) => alignment.ToLowerInvariant() switch
    {
        "middle" or "center" => A.TextAnchoringTypeValues.Center,
        "bottom" => A.TextAnchoringTypeValues.Bottom,
        _ => A.TextAnchoringTypeValues.Top
    };

    private static A.ShapeTypeValues MapShapeType(string shapeType) => shapeType.ToLowerInvariant() switch
    {
        "rectangle" or "rect" => A.ShapeTypeValues.Rectangle,
        "roundrectangle" or "roundrect" or "roundedrectangle" => A.ShapeTypeValues.RoundRectangle,
        "ellipse" or "oval" or "circle" => A.ShapeTypeValues.Ellipse,
        "triangle" => A.ShapeTypeValues.Triangle,
        "righttriangle" => A.ShapeTypeValues.RightTriangle,
        "diamond" => A.ShapeTypeValues.Diamond,
        "pentagon" => A.ShapeTypeValues.Pentagon,
        "hexagon" => A.ShapeTypeValues.Hexagon,
        "heptagon" => A.ShapeTypeValues.Heptagon,
        "octagon" => A.ShapeTypeValues.Octagon,
        "trapezoid" => A.ShapeTypeValues.Trapezoid,
        "parallelogram" => A.ShapeTypeValues.Parallelogram,
        "chevron" => A.ShapeTypeValues.Chevron,
        "homeplat" or "homeplate" => A.ShapeTypeValues.HomePlate,
        "arrow" or "rightarrow" => A.ShapeTypeValues.RightArrow,
        "leftarrow" => A.ShapeTypeValues.LeftArrow,
        "uparrow" => A.ShapeTypeValues.UpArrow,
        "downarrow" => A.ShapeTypeValues.DownArrow,
        "leftrightarrow" => A.ShapeTypeValues.LeftRightArrow,
        "updownarrow" => A.ShapeTypeValues.UpDownArrow,
        "notchedrightarrow" => A.ShapeTypeValues.NotchedRightArrow,
        "bentarrow" => A.ShapeTypeValues.BentArrow,
        "uturnarrow" => A.ShapeTypeValues.LeftUpArrow,
        "stripedrightarrow" => A.ShapeTypeValues.StripedRightArrow,
        "curvedrightarrow" => A.ShapeTypeValues.CurvedRightArrow,
        "curvedleftarrow" => A.ShapeTypeValues.CurvedLeftArrow,
        "curveduparrow" => A.ShapeTypeValues.CurvedUpArrow,
        "curveddownarrow" => A.ShapeTypeValues.CurvedDownArrow,
        "star" or "star5" => A.ShapeTypeValues.Star5,
        "star4" => A.ShapeTypeValues.Star4,
        "star6" => A.ShapeTypeValues.Star6,
        "star8" => A.ShapeTypeValues.Star8,
        "star10" => A.ShapeTypeValues.Star10,
        "star12" => A.ShapeTypeValues.Star12,
        "star16" => A.ShapeTypeValues.Star16,
        "star24" => A.ShapeTypeValues.Star24,
        "star32" => A.ShapeTypeValues.Star32,
        "heart" => A.ShapeTypeValues.Heart,
        "cloud" => A.ShapeTypeValues.Cloud,
        "lightning" or "lightningbolt" => A.ShapeTypeValues.LightningBolt,
        "sun" => A.ShapeTypeValues.Sun,
        "moon" => A.ShapeTypeValues.Moon,
        "smileyface" => A.ShapeTypeValues.SmileyFace,
        "nosmoking" => A.ShapeTypeValues.NoSmoking,
        "cross" => A.ShapeTypeValues.Plus,
        "plus" => A.ShapeTypeValues.Plus,
        "flowchartprocess" => A.ShapeTypeValues.FlowChartProcess,
        "flowchartdecision" => A.ShapeTypeValues.FlowChartDecision,
        "flowchartterminator" => A.ShapeTypeValues.FlowChartTerminator,
        "flowchartdata" or "flowchartinputoutput" => A.ShapeTypeValues.FlowChartInputOutput,
        "flowchartdocument" => A.ShapeTypeValues.FlowChartDocument,
        "flowchartmultidocument" => A.ShapeTypeValues.FlowChartMultidocument,
        "flowchartpredefinedprocess" => A.ShapeTypeValues.FlowChartPredefinedProcess,
        "flowchartpreparation" => A.ShapeTypeValues.FlowChartPreparation,
        "flowchartmanualoperation" => A.ShapeTypeValues.FlowChartManualOperation,
        "flowchartconnector" => A.ShapeTypeValues.FlowChartConnector,
        "callout1" or "callout" => A.ShapeTypeValues.Callout1,
        "callout2" => A.ShapeTypeValues.Callout2,
        "callout3" => A.ShapeTypeValues.Callout3,
        "accentcallout1" => A.ShapeTypeValues.AccentCallout1,
        "borderCallout1" or "bordercallout" => A.ShapeTypeValues.BorderCallout1,
        "wedgeroundrectcallout" or "roundedrectcallout" => A.ShapeTypeValues.WedgeRoundRectangleCallout,
        "wedgeellipsecallout" or "ovalcallout" => A.ShapeTypeValues.WedgeEllipseCallout,
        "cloudcallout" => A.ShapeTypeValues.CloudCallout,
        "leftbrace" or "brace" => A.ShapeTypeValues.LeftBrace,
        "rightbrace" => A.ShapeTypeValues.RightBrace,
        "leftbracket" or "bracket" => A.ShapeTypeValues.LeftBracket,
        "rightbracket" => A.ShapeTypeValues.RightBracket,
        "donut" or "ring" => A.ShapeTypeValues.Donut,
        "blockArc" or "arc" => A.ShapeTypeValues.BlockArc,
        "can" or "cylinder" => A.ShapeTypeValues.Can,
        "cube" => A.ShapeTypeValues.Cube,
        "ribbon" => A.ShapeTypeValues.Ribbon,
        "ribbon2" => A.ShapeTypeValues.Ribbon2,
        "foldedcorner" => A.ShapeTypeValues.FoldedCorner,
        "frame" => A.ShapeTypeValues.Frame,
        "plaque" => A.ShapeTypeValues.Plaque,
        "swoosharrow" => A.ShapeTypeValues.SwooshArrow,
        "actionbuttonhome" => A.ShapeTypeValues.ActionButtonHome,
        "line" => A.ShapeTypeValues.Line,
        "bentconnector" => A.ShapeTypeValues.BentConnector3,
        "curvedconnector" => A.ShapeTypeValues.CurvedConnector3,
        "straightconnector" => A.ShapeTypeValues.StraightConnector1,
        "mathplus" => A.ShapeTypeValues.MathPlus,
        "mathminus" => A.ShapeTypeValues.MathMinus,
        "mathmultiply" => A.ShapeTypeValues.MathMultiply,
        "mathdivide" => A.ShapeTypeValues.MathDivide,
        "mathequal" => A.ShapeTypeValues.MathEqual,
        "mathnotequal" => A.ShapeTypeValues.MathNotEqual,
        "gear6" => A.ShapeTypeValues.Gear6,
        "gear9" => A.ShapeTypeValues.Gear9,
        "roundrect" or "snip1rectangle" or "snip1rect" => A.ShapeTypeValues.Snip1Rectangle,
        "snip2rectangle" or "snip2rect" => A.ShapeTypeValues.Snip2SameRectangle,
        "round1rectangle" or "round1rect" => A.ShapeTypeValues.Round1Rectangle,
        "round2rectangle" or "round2rect" => A.ShapeTypeValues.Round2SameRectangle,
        "teardrop" => A.ShapeTypeValues.Teardrop,
        "pie" => A.ShapeTypeValues.Pie,
        "halfframe" => A.ShapeTypeValues.HalfFrame,
        "lshape" => A.ShapeTypeValues.Corner,
        "diagstripe" => A.ShapeTypeValues.DiagonalStripe,
        "chord" => A.ShapeTypeValues.Chord,
        _ => A.ShapeTypeValues.Rectangle
    };

    private static P.GraphicFrame CreateTableGraphicFrame(string[][] data, long x, long y, long width, long height)
    {
        var frameId = NextShapeId();
        var rows = data.Length;
        var cols = data.Max(r => r.Length);
        var colWidth = width / cols;
        var rowHeight = height / rows;

        var table = new A.Table(
            new A.TableProperties { FirstRow = true },
            new A.TableGrid(
                Enumerable.Range(0, cols).Select(_ => new A.GridColumn { Width = colWidth }).ToArray()
            )
        );

        for (int rowIdx = 0; rowIdx < rows; rowIdx++)
        {
            var tableRow = new A.TableRow { Height = rowHeight };
            var rowData = data[rowIdx];
            
            for (int colIdx = 0; colIdx < cols; colIdx++)
            {
                var cellText = colIdx < rowData.Length ? rowData[colIdx] : string.Empty;
                var cell = new A.TableCell(
                    new A.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(cellText)))
                    ),
                    new A.TableCellProperties(
                        new A.SolidFill(new A.SchemeColor { Val = rowIdx == 0 ? A.SchemeColorValues.Accent1 : A.SchemeColorValues.Light1 })
                    )
                );
                tableRow.Append(cell);
            }
            
            table.Append(tableRow);
        }

        return new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = frameId, Name = $"Table {frameId}" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.Transform(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = width, Cy = height }
            ),
            new A.Graphic(
                new A.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }
            )
        );
    }

    private static P.NotesSlide CreateNotesSlide()
    {
        return new P.NotesSlide(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()
                    ),
                    new P.GroupShapeProperties(),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder" },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = PlaceholderValues.Body })
                        ),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph()
                        )
                    )
                )
            ),
            new P.ColorMapOverride(new A.MasterColorMapping())
        );
    }

    private static string ExtractSlideText(SlidePart slidePart)
    {
        var texts = slidePart.Slide!.Descendants<A.Text>()
            .Select(t => t.Text)
            .Where(t => !string.IsNullOrWhiteSpace(t));
        
        return string.Join(Environment.NewLine, texts);
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
            _ => ImagePartType.Png
        };
    }

    #endregion
}
