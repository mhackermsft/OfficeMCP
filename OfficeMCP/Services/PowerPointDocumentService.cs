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

    public DocumentResult CreatePresentation(string filePath, string? title = null)
    {
        try
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using var presentation = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
            
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

            // Add first slide with title if provided
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
                "left",
                options.BackgroundColor,
                options.BorderColor
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
        TextFormatting? format = null, string alignment = "left", string? backgroundColor = null, string? borderColor = null)
    {
        var shapeId = (uint)new Random().Next(10000, 99999);
        
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"TextBox {shapeId}" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = width, Cy = height }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            ),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties { Alignment = GetTextAlignment(alignment) },
                    CreateTextRun(text, format)
                )
            )
        );

        // Add fill if background color specified
        if (!string.IsNullOrEmpty(backgroundColor))
        {
            shape.ShapeProperties!.Append(new A.SolidFill(
                new A.RgbColorModelHex { Val = backgroundColor.TrimStart('#') }
            ));
        }
        else
        {
            shape.ShapeProperties!.Append(new A.NoFill());
        }

        // Add border if specified
        if (!string.IsNullOrEmpty(borderColor))
        {
            shape.ShapeProperties!.Append(new A.Outline(
                new A.SolidFill(new A.RgbColorModelHex { Val = borderColor.TrimStart('#') })
            ));
        }

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
        var shapeId = (uint)new Random().Next(10000, 99999);
        
        var textBody = new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle()
        );

        foreach (var point in points)
        {
            var paragraph = new A.Paragraph(
                new A.ParagraphProperties(
                    new A.BulletFont { Typeface = "Arial" },
                    new A.CharacterBullet { Char = "�" }
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
        var pictureId = (uint)new Random().Next(10000, 99999);

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
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = options.WidthEmu, Cy = options.HeightEmu }
                ),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
            )
        );
    }

    private static P.Shape CreateShape(ShapeOptions options)
    {
        var shapeId = (uint)new Random().Next(10000, 99999);
        
        var shapeType = options.ShapeType.ToLowerInvariant() switch
        {
            "rectangle" or "rect" => A.ShapeTypeValues.Rectangle,
            "roundrectangle" or "roundrect" => A.ShapeTypeValues.RoundRectangle,
            "ellipse" or "oval" => A.ShapeTypeValues.Ellipse,
            "triangle" => A.ShapeTypeValues.Triangle,
            "diamond" => A.ShapeTypeValues.Diamond,
            "pentagon" => A.ShapeTypeValues.Pentagon,
            "hexagon" => A.ShapeTypeValues.Hexagon,
            "arrow" or "rightarrow" => A.ShapeTypeValues.RightArrow,
            "leftarrow" => A.ShapeTypeValues.LeftArrow,
            "star" or "star5" => A.ShapeTypeValues.Star5,
            "star4" => A.ShapeTypeValues.Star4,
            "heart" => A.ShapeTypeValues.Heart,
            _ => A.ShapeTypeValues.Rectangle
        };

        var shapeProperties = new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = options.X, Y = options.Y },
                new A.Extents { Cx = options.Width, Cy = options.Height }
            ),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = shapeType }
        );

        if (!string.IsNullOrEmpty(options.FillColor))
        {
            shapeProperties.Append(new A.SolidFill(
                new A.RgbColorModelHex { Val = options.FillColor.TrimStart('#') }
            ));
        }

        if (!string.IsNullOrEmpty(options.BorderColor))
        {
            shapeProperties.Append(new A.Outline(
                new A.SolidFill(new A.RgbColorModelHex { Val = options.BorderColor.TrimStart('#') })
            )
            { Width = (int)(options.BorderWidth * 12700) });
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = $"Shape {shapeId}" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()
            ),
            shapeProperties
        );
    }

    private static P.GraphicFrame CreateTableGraphicFrame(string[][] data, long x, long y, long width, long height)
    {
        var frameId = (uint)new Random().Next(10000, 99999);
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
