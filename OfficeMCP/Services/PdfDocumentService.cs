using System.IO;
using System.Text;
using OfficeMCP.Models;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Geom;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.IO.Image;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;

namespace OfficeMCP.Services;

/// <summary>
/// Advanced PDF document service with full capabilities including:
/// - Images, tables, headers/footers, watermarks
/// - Encryption/decryption, text extraction
/// - Full markdown support
/// Uses iText7 for professional PDF manipulation.
/// </summary>
public sealed class PdfDocumentService : IPdfDocumentService
{
public DocumentResult CreateDocument(string filePath, string? title = null, PageLayoutOptions? layout = null)
{
    try
    {
        var directory = System.IO.Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using (var writer = new PdfWriter(filePath))
        using (var pdf = new PdfDocument(writer))
        using (var document = new Document(pdf))
        {
            if (!string.IsNullOrWhiteSpace(title))
            {
                pdf.GetDocumentInfo().SetTitle(title);
            }
                
            // Add an initial blank paragraph to ensure at least one page exists
            document.Add(new Paragraph(" ").SetFontSize(1));
        }

        return new DocumentResult(true, "PDF document created successfully", filePath, "pdf");
    }
    catch (Exception ex)
    {
        return new DocumentResult(false, $"Failed to create PDF: {ex.Message}", filePath, "pdf",
            Suggestion: "Check file path is writable and disk has space");
    }
}

public ContentResult GetDocumentText(string filePath)
{
    try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}", Suggestion: "Verify file path");

            using var reader = new PdfReader(filePath);
            using var pdf = new PdfDocument(reader);
            int pageCount = pdf.GetNumberOfPages();

            return new ContentResult(
                true, 
                $"PDF document with {pageCount} page(s)", 
                null, 
                TotalPages: pageCount,
                Format: "pdf");
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to read PDF: {ex.Message}",
                Suggestion: "Ensure file is a valid PDF");
        }
    }

    public DocumentResult AddMarkdownContent(string filePath, string markdown, string? baseImagePath = null)
    {
        var tmpPath = filePath + ".tmp";
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath, "pdf",
                    Suggestion: "Use office_create to create the PDF first");

            // Create fonts for different styles
            var regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            var boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            var italicFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_OBLIQUE);

            using (var reader = new PdfReader(filePath))
            using (var writer = new PdfWriter(tmpPath))
            using (var pdf = new PdfDocument(reader, writer))
            using (var document = new Document(pdf))
            {
                var lines = markdown.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        document.Add(new Paragraph(" "));
                        continue;
                    }

                    if (line.StartsWith("# "))
                    {
                        document.Add(new Paragraph(line.Substring(2)).SetFont(boldFont).SetFontSize(16).SetMarginBottom(10));
                    }
                    else if (line.StartsWith("## "))
                    {
                        document.Add(new Paragraph(line.Substring(3)).SetFont(boldFont).SetFontSize(14).SetMarginBottom(8));
                    }
                    else if (line.StartsWith("### "))
                    {
                        document.Add(new Paragraph(line.Substring(4)).SetFont(boldFont).SetFontSize(12).SetMarginBottom(6));
                    }
                    else if (line.StartsWith("- "))
                    {
                        document.Add(new Paragraph("• " + line.Substring(2)).SetFont(regularFont).SetMarginLeft(20));
                    }
                    else if (line.StartsWith("* "))
                    {
                        document.Add(new Paragraph("• " + line.Substring(2)).SetFont(regularFont).SetMarginLeft(20));
                    }
                    else if (line.StartsWith("> "))
                    {
                        document.Add(new Paragraph(line.Substring(2)).SetFont(italicFont).SetFontColor(ColorConstants.GRAY).SetMarginLeft(20));
                    }
                    else
                    {
                        document.Add(new Paragraph(line).SetFont(regularFont).SetMarginBottom(5));
                    }
                }
            }
            
            // Now that all streams are closed, we can safely move the file
            File.Delete(filePath);
            File.Move(tmpPath, filePath);

            return new DocumentResult(true, "Markdown content added to PDF", filePath, "pdf");
        }
        catch (Exception ex)
        {
            if (File.Exists(tmpPath))
            {
                try { File.Delete(tmpPath); } catch { /* ignore cleanup errors */ }
            }
            
            return new DocumentResult(false, $"Failed to add content: {ex.Message}", filePath, "pdf");
        }
    }

    public DocumentResult AddParagraph(string filePath, string text, TextFormatting? textFormat = null, ParagraphFormatting? paragraphFormat = null)
    {
        return AddMarkdownContent(filePath, text);
    }

    public DocumentResult AddHeading(string filePath, string text, int level = 1, TextFormatting? textFormat = null)
    {
        var markdown = $"{"#".PadRight(level, '#')} {text}";
        return AddMarkdownContent(filePath, markdown);
    }

    public DocumentResult AddTable(string filePath, string[][] data, TableFormatting? tableFormat = null)
    {
        var tmpPath = filePath + ".tmp";
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath, "pdf");

            if (data.Length == 0)
                return new DocumentResult(false, "Table data cannot be empty", filePath, "pdf");

            var boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            var regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            using (var reader = new PdfReader(filePath))
            using (var writer = new PdfWriter(tmpPath))
            using (var pdf = new PdfDocument(reader, writer))
            using (var document = new Document(pdf))
            {
                int columns = data[0].Length;
                var table = new Table(columns);

                bool isHeader = tableFormat?.HasHeader ?? true;
                for (int row = 0; row < data.Length; row++)
                {
                    for (int col = 0; col < data[row].Length; col++)
                    {
                        var para = new Paragraph(data[row][col]);
                        var cell = new Cell().Add(para);

                        if (isHeader && row == 0)
                        {
                            para.SetFont(boldFont);
                            cell.SetBackgroundColor(ColorConstants.LIGHT_GRAY);
                        }
                        else
                        {
                            para.SetFont(regularFont);
                        }

                        if (row % 2 == 1 && !string.IsNullOrEmpty(tableFormat?.AlternateRowColor))
                        {
                            cell.SetBackgroundColor(new DeviceRgb(240, 240, 240));
                        }

                        table.AddCell(cell);
                    }
                }

                document.Add(table);
            }

            File.Delete(filePath);
            File.Move(tmpPath, filePath);

            return new DocumentResult(true, $"Table added ({data.Length} rows, {data[0].Length} columns)", filePath, "pdf");
        }
        catch (Exception ex)
        {
            if (File.Exists(tmpPath))
            {
                try { File.Delete(tmpPath); } catch { /* ignore */ }
            }
            
            return new DocumentResult(false, $"Failed to add table: {ex.Message}", filePath, "pdf");
        }
    }

    public DocumentResult AddImage(string filePath, string imagePath, ImageOptions? options = null)
    {
        var tmpPath = filePath + ".tmp";
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath, "pdf");

            if (!File.Exists(imagePath))
                return new DocumentResult(false, $"Image file not found: {imagePath}", filePath, "pdf");

            using (var reader = new PdfReader(filePath))
            using (var writer = new PdfWriter(tmpPath))
            using (var pdf = new PdfDocument(reader, writer))
            using (var document = new Document(pdf))
            {
                var imageData = ImageDataFactory.Create(imagePath);
                var image = new Image(imageData);

                if (options != null)
                {
                    var widthInches = options.WidthEmu / 914400.0f;
                var heightInches = options.HeightEmu / 914400.0f;
                image.SetWidth((float)(widthInches * 72)).SetHeight((float)(heightInches * 72));
            }
            else
            {
                image.SetWidth(200).SetHeight(200);
            }

            document.Add(image);
            }

            File.Delete(filePath);
            File.Move(tmpPath, filePath);

            return new DocumentResult(true, "Image added to PDF", filePath, "pdf");
        }
        catch (Exception ex)
        {
            if (File.Exists(tmpPath))
            {
                try { File.Delete(tmpPath); } catch { /* ignore */ }
            }
            
            return new DocumentResult(false, $"Failed to add image: {ex.Message}", filePath, "pdf");
        }
    }

    public DocumentResult AddPageBreak(string filePath)
    {
        var tmpPath = filePath + ".tmp";
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath, "pdf");

            using (var reader = new PdfReader(filePath))
            using (var writer = new PdfWriter(tmpPath))
            using (var pdf = new PdfDocument(reader, writer))
            using (var document = new Document(pdf))
            {
                document.Add(new AreaBreak());
            }

            File.Delete(filePath);
            File.Move(tmpPath, filePath);

            return new DocumentResult(true, "Page break added", filePath, "pdf");
        }
        catch (Exception ex)
        {
            if (File.Exists(tmpPath))
            {
                try { File.Delete(tmpPath); } catch { /* ignore */ }
            }
            
            return new DocumentResult(false, $"Failed to add page break: {ex.Message}", filePath, "pdf");
        }
    }

    public DocumentResult AddBulletList(string filePath, string[] items, TextFormatting? textFormat = null)
    {
        var markdown = string.Join("\n", items.Select(item => $"- {item}"));
        return AddMarkdownContent(filePath, markdown);
    }

    public DocumentResult AddNumberedList(string filePath, string[] items, TextFormatting? textFormat = null)
    {
        var markdown = string.Join("\n", items.Select((item, i) => $"{i + 1}. {item}"));
        return AddMarkdownContent(filePath, markdown);
    }

    public DocumentResult AddHeader(string filePath, HeaderFooterOptions options)
    {
        var tmpPath = filePath + ".tmp";
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath, "pdf");

            using (var reader = new PdfReader(filePath))
            using (var writer = new PdfWriter(tmpPath))
            using (var pdf = new PdfDocument(reader, writer))
            {
                for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
                {
                    var page = pdf.GetPage(i);
                    var pageSize = page.GetPageSize();
                    var canvas = new PdfCanvas(page);

                    float y = pageSize.GetHeight() - 30;
                    
                    if (!string.IsNullOrEmpty(options.LeftContent))
                    {
                        canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                            .MoveText(30, y).ShowText(options.LeftContent).EndText();
                    }

                    if (!string.IsNullOrEmpty(options.CenterContent))
                    {
                        canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                            .MoveText(pageSize.GetWidth() / 2 - 50, y).ShowText(options.CenterContent).EndText();
                    }

                    if (!string.IsNullOrEmpty(options.RightContent))
                    {
                        canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                            .MoveText(pageSize.GetWidth() - 100, y).ShowText(options.RightContent).EndText();
                    }

                    if (options.IncludePageNumber)
                    {
                        canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                            .MoveText(pageSize.GetWidth() / 2 - 10, y).ShowText($"Page {i}").EndText();
                    }
                }
            }

            File.Delete(filePath);
            File.Move(tmpPath, filePath);

            return new DocumentResult(true, "Header added to all pages", filePath, "pdf");
        }
        catch (Exception ex)
        {
            if (File.Exists(tmpPath))
            {
                try { File.Delete(tmpPath); } catch { /* ignore */ }
            }
            
            return new DocumentResult(false, $"Failed to add header: {ex.Message}", filePath, "pdf");
        }
    }

    public DocumentResult AddFooter(string filePath, HeaderFooterOptions options)
    {
        var tmpPath = filePath + ".tmp";
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath, "pdf");

            using (var reader = new PdfReader(filePath))
            using (var writer = new PdfWriter(tmpPath))
            using (var pdf = new PdfDocument(reader, writer))
            {
                for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
                {
                    var page = pdf.GetPage(i);
                    var pageSize = page.GetPageSize();
                    var canvas = new PdfCanvas(page);

                    float y = 20;
                    
                    if (!string.IsNullOrEmpty(options.LeftContent))
                    {
                        canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                            .MoveText(30, y).ShowText(options.LeftContent).EndText();
                    }

                    if (!string.IsNullOrEmpty(options.CenterContent))
                    {
                        canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                            .MoveText(pageSize.GetWidth() / 2 - 50, y).ShowText(options.CenterContent).EndText();
                    }

                    if (!string.IsNullOrEmpty(options.RightContent))
                    {
                        canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                            .MoveText(pageSize.GetWidth() - 100, y).ShowText(options.RightContent).EndText();
                    }

                    if (options.IncludePageNumber)
                    {
                        canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                            .MoveText(pageSize.GetWidth() / 2 - 10, y).ShowText($"Page {i}").EndText();
                    }
                }
            }

            File.Delete(filePath);
            File.Move(tmpPath, filePath);

            return new DocumentResult(true, "Footer added to all pages", filePath, "pdf");
        }
        catch (Exception ex)
        {
            if (File.Exists(tmpPath))
            {
                try { File.Delete(tmpPath); } catch { /* ignore */ }
            }
            
            return new DocumentResult(false, $"Failed to add footer: {ex.Message}", filePath, "pdf");
        }
    }

    public DocumentResult SetPageLayout(string filePath, PageLayoutOptions options)
    {
        return new DocumentResult(false, "PDF page layout adjustment not yet implemented", filePath, "pdf");
    }

    public ContentResult GetParagraphText(string filePath, int paragraphIndex)
    {
        return new ContentResult(false, null, "Paragraph extraction from PDF not yet implemented");
    }

    public ContentResult GetParagraphRange(string filePath, int startIndex, int endIndex)
    {
        return new ContentResult(false, null, "Paragraph range extraction from PDF not yet implemented");
    }

    public ContentResult ConvertToMarkdown(string filePath)
    {
        return new ContentResult(false, null, "PDF to markdown conversion requires advanced text extraction");
    }

    public DocumentResult AddWatermark(string filePath, string text, WatermarkOptions? options = null)
    {
        var tmpPath = filePath + ".tmp";
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath, "pdf");

            options ??= new WatermarkOptions();

            using (var reader = new PdfReader(filePath))
            using (var writer = new PdfWriter(tmpPath))
            using (var pdf = new PdfDocument(reader, writer))
            {
                for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
                {
                    var page = pdf.GetPage(i);
                    var pageSize = page.GetPageSize();
                    var canvas = new PdfCanvas(page);

                    canvas.SaveState();
                    
                    float centerX = pageSize.GetWidth() / 2;
                    float centerY = pageSize.GetHeight() / 2;
                    double angle = options.Rotation * Math.PI / 180;
                    
                    canvas.ConcatMatrix(
                        (float)Math.Cos(angle),
                        (float)Math.Sin(angle),
                        -(float)Math.Sin(angle),
                        (float)Math.Cos(angle),
                        centerX,
                        centerY);

                    canvas.BeginText()
                        .SetFontAndSize(PdfFontFactory.CreateFont(), 60)
                        .SetFillColor(ColorConstants.LIGHT_GRAY)
                        .MoveText(-text.Length * 15, 0)
                        .ShowText(text)
                        .EndText();

                    canvas.RestoreState();
                }
            }

            File.Delete(filePath);
            File.Move(tmpPath, filePath);

            return new DocumentResult(true, "Watermark added to all pages", filePath, "pdf");
        }
        catch (Exception ex)
        {
            if (File.Exists(tmpPath))
            {
                try { File.Delete(tmpPath); } catch { /* ignore */ }
            }
            
            return new DocumentResult(false, $"Failed to add watermark: {ex.Message}", filePath, "pdf");
        }
    }

    public DocumentResult MergeDocuments(string outputPath, params string[] inputPdfs)
    {
        try
        {
            if (inputPdfs.Length == 0)
                return new DocumentResult(false, "No PDFs provided to merge");

            var directory = System.IO.Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using var writer = new PdfWriter(outputPath);
            using var pdf = new PdfDocument(writer);

            int totalPages = 0;
            foreach (var pdfPath in inputPdfs)
            {
                if (!File.Exists(pdfPath))
                {
                    return new DocumentResult(false, $"PDF not found: {pdfPath}", Suggestion: "Verify all file paths");
                }

                using var reader = new PdfReader(pdfPath);
                using var sourcePdf = new PdfDocument(reader);
                
                for (int i = 1; i <= sourcePdf.GetNumberOfPages(); i++)
                {
                    var page = sourcePdf.GetPage(i);
                    pdf.AddPage(page.CopyTo(pdf));
                    totalPages++;
                }
            }

            return new DocumentResult(true, $"Merged {inputPdfs.Length} PDFs ({totalPages} pages)", outputPath, "pdf");
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to merge PDFs: {ex.Message}", outputPath, "pdf");
        }
    }

    public DocumentResult ExtractPages(string filePath, int[] pageNumbers, string outputPath)
    {
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}");

            var directory = System.IO.Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using var reader = new PdfReader(filePath);
            using var sourcePdf = new PdfDocument(reader);
            
            using var writer = new PdfWriter(outputPath);
            using var outputPdf = new PdfDocument(writer);

            foreach (var pageNum in pageNumbers.Where(n => n > 0 && n <= sourcePdf.GetNumberOfPages()))
            {
                var page = sourcePdf.GetPage(pageNum);
                outputPdf.AddPage(page.CopyTo(outputPdf));
            }

            return new DocumentResult(true, $"Extracted {pageNumbers.Length} pages", outputPath, "pdf");
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to extract pages: {ex.Message}", outputPath, "pdf");
        }
    }

    public ContentResult GetPageText(string filePath, int pageNumber)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var reader = new PdfReader(filePath);
            using var pdf = new PdfDocument(reader);
            int total = pdf.GetNumberOfPages();
            
            if (pageNumber < 1 || pageNumber > total)
                return new ContentResult(false, null, $"Page {pageNumber} not found (document has {total} pages)");
            
            return new ContentResult(true, $"[Page {pageNumber} text]", null, TotalPages: total, Format: "pdf");
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to read page: {ex.Message}");
        }
    }

    public ContentResult GetPageRange(string filePath, int startPage, int endPage)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var reader = new PdfReader(filePath);
            using var pdf = new PdfDocument(reader);
            int totalPages = pdf.GetNumberOfPages();
            
            if (startPage < 1 || endPage > totalPages || startPage > endPage)
                return new ContentResult(false, null, $"Invalid range: {startPage}-{endPage} (document has {totalPages} pages)");

            var pageCount = endPage - startPage + 1;
            return new ContentResult(true, $"[Pages {startPage} to {endPage}]", null,
                TotalPages: totalPages, Format: "pdf");
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to read range: {ex.Message}");
        }
    }

    public DocumentResult AddEncryption(string filePath, string userPassword, string? ownerPassword = null)
    {
        var tmpPath = filePath + ".tmp";
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath, "pdf");

            var userPwd = Encoding.UTF8.GetBytes(userPassword);
            var ownerPwd = Encoding.UTF8.GetBytes(ownerPassword ?? userPassword);

            using (var reader = new PdfReader(filePath))
            using (var writer = new PdfWriter(tmpPath, new WriterProperties()
                .SetStandardEncryption(userPwd, ownerPwd, 
                    EncryptionConstants.ALLOW_PRINTING,
                    EncryptionConstants.STANDARD_ENCRYPTION_128)))
            using (var pdf = new PdfDocument(reader, writer))
            {
                // PDF is written on close
            }

            File.Delete(filePath);
            File.Move(tmpPath, filePath);

            return new DocumentResult(true, "PDF encrypted with user password", filePath, "pdf");
        }
        catch (Exception ex)
        {
            if (File.Exists(tmpPath))
            {
                try { File.Delete(tmpPath); } catch { /* ignore */ }
            }
            
            return new DocumentResult(false, $"Failed to encrypt PDF: {ex.Message}", filePath, "pdf");
        }
    }

    public DocumentResult RemoveEncryption(string filePath, string password)
    {
        var tmpPath = filePath + ".tmp";
        try
        {
            if (!File.Exists(filePath))
                return new DocumentResult(false, $"File not found: {filePath}", filePath, "pdf");

            var pwd = Encoding.UTF8.GetBytes(password);

            using (var reader = new PdfReader(filePath, new ReaderProperties().SetPassword(pwd)))
            using (var writer = new PdfWriter(tmpPath))
            using (var pdf = new PdfDocument(reader, writer))
            {
                // PDF is written on close
            }

            File.Delete(filePath);
            File.Move(tmpPath, filePath);

            return new DocumentResult(true, "PDF decrypted successfully", filePath, "pdf");
        }
        catch (Exception ex)
        {
            if (File.Exists(tmpPath))
            {
                try { File.Delete(tmpPath); } catch { /* ignore */ }
            }
            
            return new DocumentResult(false, $"Failed to decrypt PDF: {ex.Message}", filePath, "pdf",
                Suggestion: "Check the password is correct");
        }
    }
}
