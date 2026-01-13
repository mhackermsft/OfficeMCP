using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeMCP.Models;
using System.Text;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeMCP.Services;

/// <summary>
/// Service for creating and manipulating Excel workbooks using OpenXML.
/// </summary>
public sealed partial class ExcelDocumentService : IExcelDocumentService
{
    public DocumentResult CreateWorkbook(string filePath, string? sheetName = null)
    {
        try
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using var spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            
            var workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = sheetName ?? "Sheet1"
            };
            sheets.Append(sheet);

            // Add default stylesheet
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateDefaultStylesheet();

            spreadsheet.Save();
            return new DocumentResult(true, $"Workbook created successfully at {filePath}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to create workbook: {ex.Message}");
        }
    }

    public DocumentResult AddSheet(string filePath, string sheetName)
    {
        try
        {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var workbookPart = spreadsheet.WorkbookPart;
            
            if (workbookPart == null)
                return new DocumentResult(false, "Workbook part not found");

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            var sheetId = sheets?.Elements<Sheet>().Max(s => s.SheetId?.Value ?? 0) + 1 ?? 1;

            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)sheetId,
                Name = sheetName
            };
            sheets?.Append(sheet);

            spreadsheet.Save();
            return new DocumentResult(true, $"Sheet '{sheetName}' added successfully", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add sheet: {ex.Message}");
        }
    }

    public DocumentResult SetCellValue(string filePath, string sheetName, string cellReference, string value, ExcelCellFormatting? formatting = null)
    {
        try
        {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                return new DocumentResult(false, "Sheet data not found");

            var cell = GetOrCreateCell(sheetData, cellReference);
            SetCellValueInternal(cell, value);

            if (formatting != null)
            {
                ApplyCellFormatting(spreadsheet, cell, formatting);
            }

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Cell {cellReference} set to '{value}'", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to set cell value: {ex.Message}");
        }
    }

    public DocumentResult SetRangeValues(string filePath, string sheetName, string startCell, string[][] values)
    {
        try
        {
            if (values.Length == 0)
                return new DocumentResult(false, "Values array cannot be empty");

            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                return new DocumentResult(false, "Sheet data not found");

            var (startCol, startRow) = ParseCellReference(startCell);
            
            for (int rowOffset = 0; rowOffset < values.Length; rowOffset++)
            {
                var rowData = values[rowOffset];
                for (int colOffset = 0; colOffset < rowData.Length; colOffset++)
                {
                    var cellRef = GetCellReference(startCol + colOffset, startRow + rowOffset);
                    var cell = GetOrCreateCell(sheetData, cellRef);
                    SetCellValueInternal(cell, rowData[colOffset]);
                }
            }

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Range starting at {startCell} populated with {values.Length} rows", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to set range values: {ex.Message}");
        }
    }

    public DocumentResult AddTable(string filePath, string sheetName, string startCell, string[][] data, bool hasHeaders = true)
    {
        try
        {
            if (data.Length == 0)
                return new DocumentResult(false, "Table data cannot be empty");

            // First, set all the values
            var setResult = SetRangeValues(filePath, sheetName, startCell, data);
            if (!setResult.Success)
                return setResult;

            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            // Calculate end cell
            var (startCol, startRow) = ParseCellReference(startCell);
            var maxCols = data.Max(row => row.Length);
            var endCell = GetCellReference(startCol + maxCols - 1, startRow + data.Length - 1);
            var rangeRef = $"{startCell}:{endCell}";

            // Add table definition
            var tableDefPart = worksheetPart.AddNewPart<TableDefinitionPart>();
            var tableId = (uint)(worksheetPart.TableDefinitionParts.Count());
            
            var tableColumns = new TableColumns { Count = (uint)maxCols };
            for (int i = 0; i < maxCols; i++)
            {
                var colName = hasHeaders && data[0].Length > i ? data[0][i] : $"Column{i + 1}";
                tableColumns.Append(new TableColumn { Id = (uint)(i + 1), Name = colName });
            }

            var table = new Table
            {
                Id = tableId,
                Name = $"Table{tableId}",
                DisplayName = $"Table{tableId}",
                Reference = rangeRef,
                TotalsRowShown = false
            };

            var autoFilter = new AutoFilter { Reference = rangeRef };
            var tableStyleInfo = new TableStyleInfo
            {
                Name = "TableStyleMedium2",
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            };

            table.Append(autoFilter);
            table.Append(tableColumns);
            table.Append(tableStyleInfo);
            tableDefPart.Table = table;

            // Add table parts reference to worksheet
            var tableParts = worksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
            if (tableParts == null)
            {
                tableParts = new TableParts();
                worksheetPart.Worksheet.Append(tableParts);
            }
            tableParts.Append(new TablePart { Id = worksheetPart.GetIdOfPart(tableDefPart) });
            tableParts.Count = (uint)tableParts.Elements<TablePart>().Count();

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Table with {data.Length} rows added at {startCell}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add table: {ex.Message}");
        }
    }

    public DocumentResult AddImage(string filePath, string sheetName, string imagePath, string cellReference, ImageOptions? options = null)
    {
        try
        {
            if (!File.Exists(imagePath))
                return new DocumentResult(false, $"Image file not found: {imagePath}");

            options ??= new ImageOptions();

            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null)
            {
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                
                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            var imagePart = drawingsPart.AddImagePart(GetImagePartType(imagePath));
            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }

            var (col, row) = ParseCellReference(cellReference);
            AddImageToDrawing(drawingsPart, imagePart, col, row, options);

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Image added at cell {cellReference}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add image: {ex.Message}");
        }
    }

    public DocumentResult MergeCells(string filePath, string sheetName, string startCell, string endCell)
    {
        try
        {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var worksheet = worksheetPart.Worksheet;
            var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault();
            
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
            }

            var mergeCell = new MergeCell { Reference = $"{startCell}:{endCell}" };
            mergeCells.Append(mergeCell);
            mergeCells.Count = (uint)mergeCells.Elements<MergeCell>().Count();

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Cells {startCell}:{endCell} merged", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to merge cells: {ex.Message}");
        }
    }

    public DocumentResult SetColumnWidth(string filePath, string sheetName, int columnIndex, double width)
    {
        try
        {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var worksheet = worksheetPart.Worksheet;
            var columns = worksheet.Elements<Columns>().FirstOrDefault();
            
            if (columns == null)
            {
                columns = new Columns();
                worksheet.InsertBefore(columns, worksheet.Elements<SheetData>().First());
            }

            var column = columns.Elements<Column>().FirstOrDefault(c => 
                c.Min?.Value <= columnIndex && c.Max?.Value >= columnIndex);
            
            if (column != null)
            {
                column.Width = width;
                column.CustomWidth = true;
            }
            else
            {
                columns.Append(new Column
                {
                    Min = (uint)columnIndex,
                    Max = (uint)columnIndex,
                    Width = width,
                    CustomWidth = true
                });
            }

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Column {columnIndex} width set to {width}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to set column width: {ex.Message}");
        }
    }

    public DocumentResult SetRowHeight(string filePath, string sheetName, int rowIndex, double height)
    {
        try
        {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var row = sheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
            
            if (row == null)
            {
                row = new Row { RowIndex = (uint)rowIndex };
                sheetData?.Append(row);
            }

            row.Height = height;
            row.CustomHeight = true;

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Row {rowIndex} height set to {height}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to set row height: {ex.Message}");
        }
    }

    public DocumentResult AddFormula(string filePath, string sheetName, string cellReference, string formula)
    {
        try
        {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                return new DocumentResult(false, "Sheet data not found");

            var cell = GetOrCreateCell(sheetData, cellReference);
            cell.CellFormula = new CellFormula(formula);
            cell.CellValue = null; // Clear any existing value, formula will calculate

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Formula set in cell {cellReference}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to add formula: {ex.Message}");
        }
    }

    public DocumentResult AutoFitColumn(string filePath, string sheetName, int columnIndex)
    {
        // OpenXML doesn't have auto-fit capability - set a reasonable default width
        return SetColumnWidth(filePath, sheetName, columnIndex, 12.0);
    }

    public ContentResult GetCellValue(string filePath, string sheetName, string cellReference)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new ContentResult(false, null, $"Sheet '{sheetName}' not found");

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var cell = sheetData?.Descendants<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellReference);

            if (cell == null)
                return new ContentResult(true, string.Empty);

            var value = GetCellValueAsString(spreadsheet, cell);
            return new ContentResult(true, value);
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get cell value: {ex.Message}");
        }
    }

    public ContentResult GetRangeValues(string filePath, string sheetName, string startCell, string endCell)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new ContentResult(false, null, $"Sheet '{sheetName}' not found");

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                return new ContentResult(false, null, "Sheet data not found");

            var (startCol, startRow) = ParseCellReference(startCell);
            var (endCol, endRow) = ParseCellReference(endCell);

            var sb = new StringBuilder();
            for (int row = startRow; row <= endRow; row++)
            {
                var rowValues = new List<string>();
                for (int col = startCol; col <= endCol; col++)
                {
                    var cellRef = GetCellReference(col, row);
                    var cell = sheetData.Descendants<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRef);
                    rowValues.Add(cell != null ? GetCellValueAsString(spreadsheet, cell) : string.Empty);
                }
                sb.AppendLine(string.Join("\t", rowValues));
            }

            return new ContentResult(true, sb.ToString().TrimEnd());
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get range values: {ex.Message}");
        }
    }

    public ContentResult GetSheetText(string filePath, string sheetName)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new ContentResult(false, null, $"Sheet '{sheetName}' not found");

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                return new ContentResult(false, null, "Sheet data not found");

            var sb = new StringBuilder();
            foreach (var row in sheetData.Elements<Row>().OrderBy(r => r.RowIndex?.Value ?? 0))
            {
                var rowValues = new List<string>();
                foreach (var cell in row.Elements<Cell>())
                {
                    rowValues.Add(GetCellValueAsString(spreadsheet, cell));
                }
                if (rowValues.Count != 0)
                {
                    sb.AppendLine(string.Join("\t", rowValues));
                }
            }

            return new ContentResult(true, sb.ToString().TrimEnd());
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get sheet text: {ex.Message}");
        }
    }

    public ContentResult GetAllSheetsText(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            using var spreadsheet = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = spreadsheet.WorkbookPart;
            
            if (workbookPart == null)
                return new ContentResult(false, null, "Workbook part not found");

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            var sb = new StringBuilder();

            foreach (var sheet in sheets?.Elements<Sheet>() ?? [])
            {
                sb.AppendLine($"=== Sheet: {sheet.Name} ===");
                var result = GetSheetText(filePath, sheet.Name?.Value ?? string.Empty);
                if (result.Success)
                {
                    sb.AppendLine(result.Content);
                }
                sb.AppendLine();
            }

            return new ContentResult(true, sb.ToString().TrimEnd());
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get workbook text: {ex.Message}");
        }
    }

    public DocumentResult DeleteSheet(string filePath, string sheetName)
    {
        try
        {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var workbookPart = spreadsheet.WorkbookPart;
            
            if (workbookPart == null)
                return new DocumentResult(false, "Workbook part not found");

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            var sheet = sheets?.Elements<Sheet>().FirstOrDefault(s => s.Name?.Value == sheetName);
            
            if (sheet == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var relationshipId = sheet.Id?.Value;
            if (relationshipId != null)
            {
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                workbookPart.DeletePart(worksheetPart);
            }

            sheet.Remove();
            spreadsheet.Save();
            return new DocumentResult(true, $"Sheet '{sheetName}' deleted", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to delete sheet: {ex.Message}");
        }
    }

    public DocumentResult RenameSheet(string filePath, string oldName, string newName)
    {
        try
        {
            using var spreadsheet = SpreadsheetDocument.Open(filePath, true);
            var workbookPart = spreadsheet.WorkbookPart;
            
            if (workbookPart == null)
                return new DocumentResult(false, "Workbook part not found");

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            var sheet = sheets?.Elements<Sheet>().FirstOrDefault(s => s.Name?.Value == oldName);
            
            if (sheet == null)
                return new DocumentResult(false, $"Sheet '{oldName}' not found");

            sheet.Name = newName;
            spreadsheet.Save();
            return new DocumentResult(true, $"Sheet renamed from '{oldName}' to '{newName}'", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to rename sheet: {ex.Message}");
        }
    }

    #region Private Helper Methods

    private static WorksheetPart? GetWorksheetPart(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook.GetFirstChild<Sheets>()?.Elements<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        
        if (sheet?.Id?.Value == null) return null;
        
        return (WorksheetPart)workbookPart!.GetPartById(sheet.Id.Value);
    }

    private static Cell GetOrCreateCell(SheetData sheetData, string cellReference)
    {
        var (col, rowIndex) = ParseCellReference(cellReference);
        
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = (uint)rowIndex };
            
            // Insert row in correct position
            var refRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value > rowIndex);
            if (refRow != null)
            {
                sheetData.InsertBefore(row, refRow);
            }
            else
            {
                sheetData.Append(row);
            }
        }

        var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellReference);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellReference };
            
            // Insert cell in correct position
            var refCell = row.Elements<Cell>().FirstOrDefault(c => 
            {
                var (refCol, _) = ParseCellReference(c.CellReference?.Value ?? "A1");
                return refCol > col;
            });
            
            if (refCell != null)
            {
                row.InsertBefore(cell, refCell);
            }
            else
            {
                row.Append(cell);
            }
        }

        return cell;
    }

    private static void SetCellValueInternal(Cell cell, string value)
    {
        if (double.TryParse(value, out var numValue))
        {
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(numValue.ToString());
        }
        else if (bool.TryParse(value, out var boolValue))
        {
            cell.DataType = CellValues.Boolean;
            cell.CellValue = new CellValue(boolValue ? "1" : "0");
        }
        else
        {
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(value);
        }
    }

    private static string GetCellValueAsString(SpreadsheetDocument document, Cell cell)
    {
        var value = cell.CellValue?.Text ?? string.Empty;
        
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var stringTable = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
            if (stringTable != null && int.TryParse(value, out var index))
            {
                value = stringTable.ElementAt(index).InnerText;
            }
        }

        return value;
    }

    private static (int Column, int Row) ParseCellReference(string cellReference)
    {
        var match = CellReferenceRegex().Match(cellReference.ToUpperInvariant());
        if (!match.Success)
            throw new ArgumentException($"Invalid cell reference: {cellReference}");

        var colStr = match.Groups[1].Value;
        var rowStr = match.Groups[2].Value;

        int col = 0;
        foreach (char c in colStr)
        {
            col = col * 26 + (c - 'A' + 1);
        }

        return (col, int.Parse(rowStr));
    }

    private static string GetCellReference(int column, int row)
    {
        var colStr = string.Empty;
        while (column > 0)
        {
            column--;
            colStr = (char)('A' + column % 26) + colStr;
            column /= 26;
        }
        return $"{colStr}{row}";
    }

    [GeneratedRegex(@"^([A-Z]+)(\d+)$")]
    private static partial Regex CellReferenceRegex();

    private static Stylesheet CreateDefaultStylesheet()
    {
        return new Stylesheet(
            new Fonts(
                new Font(
                    new FontSize { Val = 11 },
                    new FontName { Val = "Calibri" }
                ),
                new Font(
                    new Bold(),
                    new FontSize { Val = 11 },
                    new FontName { Val = "Calibri" }
                )
            ),
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            ),
            new Borders(
                new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()
                )
            ),
            new CellFormats(
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0 },
                new CellFormat { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true }
            )
        );
    }

    private static void ApplyCellFormatting(SpreadsheetDocument document, Cell cell, ExcelCellFormatting formatting)
    {
        // Simplified formatting - full implementation would require more complex stylesheet management
        if (formatting.Bold)
        {
            cell.StyleIndex = 1; // Bold style from default stylesheet
        }
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

    private static void AddImageToDrawing(DrawingsPart drawingsPart, ImagePart imagePart, int col, int row, ImageOptions options)
    {
        var worksheetDrawing = drawingsPart.WorksheetDrawing;
        var relationshipId = drawingsPart.GetIdOfPart(imagePart);

        var twoCellAnchor = new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId((col - 1).ToString()),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((row - 1).ToString()),
                new Xdr.RowOffset("0")
            ),
            new Xdr.ToMarker(
                new Xdr.ColumnId((col + 3).ToString()),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((row + 10).ToString()),
                new Xdr.RowOffset("0")
            ),
            new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = 1, Name = "Picture 1", Description = options.AltText ?? string.Empty },
                    new Xdr.NonVisualPictureDrawingProperties()
                ),
                new Xdr.BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())
                ),
                new Xdr.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = options.WidthEmu, Cy = options.HeightEmu }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            ),
            new Xdr.ClientData()
        );

        worksheetDrawing.Append(twoCellAnchor);
    }

    #endregion
}
