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
    // OpenSettings for better compatibility with various Excel file formats
    private static readonly OpenSettings DefaultOpenSettings = new()
    {
        AutoSave = false,
        MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
            MarkupCompatibilityProcessMode.ProcessAllParts,
            DocumentFormat.OpenXml.FileFormatVersions.Microsoft365)
    };

    /// <summary>
    /// Opens a spreadsheet document with proper error handling for various file formats.
    /// </summary>
    private static SpreadsheetDocument OpenSpreadsheet(Stream stream, bool isEditable)
    {
        try
        {
            return SpreadsheetDocument.Open(stream, isEditable, DefaultOpenSettings);
        }
        catch (OpenXmlPackageException)
        {
            // Try without custom settings as fallback
            stream.Position = 0;
            return SpreadsheetDocument.Open(stream, isEditable);
        }
    }

    /// <summary>
    /// Opens a spreadsheet document from a file path with proper error handling.
    /// </summary>
    private static SpreadsheetDocument OpenSpreadsheet(string filePath, bool isEditable)
    {
        try
        {
            return SpreadsheetDocument.Open(filePath, isEditable, DefaultOpenSettings);
        }
        catch (OpenXmlPackageException)
        {
            // Try without custom settings as fallback
            return SpreadsheetDocument.Open(filePath, isEditable);
        }
    }

    /// <summary>
    /// Checks if a file is protected/encrypted and returns an appropriate error message.
    /// </summary>
    private static ContentResult? CheckFileProtectionForRead(string filePath)
    {
        var protectionInfo = OfficeFileProtectionDetector.CheckFileProtection(filePath);
        
        if (protectionInfo.IsEncrypted)
        {
            var message = OfficeFileProtectionDetector.GetProtectionErrorMessage(protectionInfo, filePath);
            return new ContentResult(false, null, message);
        }

        if (!protectionInfo.IsValidOfficeFormat && protectionInfo.ErrorMessage != null)
        {
            return new ContentResult(false, null, protectionInfo.ErrorMessage);
        }

        return null; // File is OK to open
    }

    /// <summary>
    /// Checks if a file is protected/encrypted and returns an appropriate error message for write operations.
    /// </summary>
    private static DocumentResult? CheckFileProtectionForWrite(string filePath)
    {
        var protectionInfo = OfficeFileProtectionDetector.CheckFileProtection(filePath);
        
        if (protectionInfo.IsEncrypted)
        {
            var message = OfficeFileProtectionDetector.GetProtectionErrorMessage(protectionInfo, filePath);
            return new DocumentResult(false, message);
        }

        if (!protectionInfo.IsValidOfficeFormat && protectionInfo.ErrorMessage != null)
        {
            return new DocumentResult(false, protectionInfo.ErrorMessage);
        }

        return null; // File is OK to open
    }

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
            using var spreadsheet = OpenSpreadsheet(filePath, true);
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
        catch (OpenXmlPackageException ex)
        {
            return new DocumentResult(false, $"Failed to open Excel file (may be corrupted, password-protected, or in an unsupported format): {ex.Message}");
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
            using var spreadsheet = OpenSpreadsheet(filePath, true);
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
        catch (OpenXmlPackageException ex)
        {
            return new DocumentResult(false, $"Failed to open Excel file (may be corrupted, password-protected, or in an unsupported format): {ex.Message}");
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

            using var spreadsheet = OpenSpreadsheet(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);
            
            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                return new DocumentResult(false, "Sheet data not found");

            var (startCol, startRow) = ParseCellReference(startCell);
            var maxCols = values.Max(row => row.Length);
            
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

            // Calculate end cell for the data range
            var endCol = startCol + maxCols - 1;
            var endRow = startRow + values.Length - 1;
            var endCell = GetCellReference(endCol, endRow);

            // Auto-expand any tables that overlap with the new data
            ExpandTablesForRange(worksheetPart, startCol, startRow, endCol, endRow);

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Range starting at {startCell} populated with {values.Length} rows", filePath);
        }
        catch (OpenXmlPackageException ex)
        {
            return new DocumentResult(false, $"Failed to open Excel file (may be corrupted, password-protected, or in an unsupported format): {ex.Message}");
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

            using var spreadsheet = OpenSpreadsheet(filePath, true);
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

            using var spreadsheet = OpenSpreadsheet(filePath, true);
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
            using var spreadsheet = OpenSpreadsheet(filePath, true);
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
            using var spreadsheet = OpenSpreadsheet(filePath, true);
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
            using var spreadsheet = OpenSpreadsheet(filePath, true);
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
            using var spreadsheet = OpenSpreadsheet(filePath, true);
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

    public DocumentResult ResizeTableToIncludeRange(string filePath, string sheetName, string startCell, string endCell)
    {
        try
        {
            using var spreadsheet = OpenSpreadsheet(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);

            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var (startCol, startRow) = ParseCellReference(startCell);
            var (endCol, endRow) = ParseCellReference(endCell);

            // Find tables that overlap with the given range
            var tableResized = false;
            foreach (var tableDefPart in worksheetPart.TableDefinitionParts)
            {
                var table = tableDefPart.Table;
                if (table?.Reference?.Value == null) continue;

                var tableRange = table.Reference.Value;
                var parts = tableRange.Split(':');
                if (parts.Length != 2) continue;

                var (tableStartCol, tableStartRow) = ParseCellReference(parts[0]);
                var (tableEndCol, tableEndRow) = ParseCellReference(parts[1]);

                // Check if the new range overlaps or extends the table
                // Table should be expanded if new data starts at or near table columns
                if (startCol >= tableStartCol && startCol <= tableEndCol + 1)
                {
                    // Expand table to include new rows/columns
                    var newEndCol = Math.Max(tableEndCol, endCol);
                    var newEndRow = Math.Max(tableEndRow, endRow);
                    var newStartCol = Math.Min(tableStartCol, startCol);
                    var newStartRow = Math.Min(tableStartRow, startRow);

                    var newStartCellRef = GetCellReference(newStartCol, newStartRow);
                    var newEndCellRef = GetCellReference(newEndCol, newEndRow);
                    var newRange = $"{newStartCellRef}:{newEndCellRef}";

                    table.Reference = newRange;

                    // Update AutoFilter if present
                    var autoFilter = table.AutoFilter;
                    if (autoFilter != null)
                    {
                        autoFilter.Reference = newRange;
                    }

                    // Update table columns if columns were added
                    var tableColumns = table.TableColumns;
                    if (tableColumns != null)
                    {
                        var currentColCount = (int)(tableColumns.Count?.Value ?? 0);
                        var newColCount = newEndCol - newStartCol + 1;

                        for (int i = currentColCount; i < newColCount; i++)
                        {
                            tableColumns.Append(new TableColumn { Id = (uint)(i + 1), Name = $"Column{i + 1}" });
                        }
                        tableColumns.Count = (uint)newColCount;
                    }

                    tableDefPart.Table.Save();
                    tableResized = true;
                }
            }

            if (tableResized)
            {
                worksheetPart.Worksheet.Save();
                spreadsheet.Save();
                return new DocumentResult(true, $"Table(s) resized to include range {startCell}:{endCell}", filePath);
            }

            return new DocumentResult(true, "No tables found that needed resizing", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to resize table: {ex.Message}");
        }
    }

    public DocumentResult FormatCellRange(string filePath, string sheetName, string startCell, string endCell, ExcelCellFormatting formatting)
    {
        try
        {
            using var spreadsheet = OpenSpreadsheet(filePath, true);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);

            if (worksheetPart == null)
                return new DocumentResult(false, $"Sheet '{sheetName}' not found");

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                return new DocumentResult(false, "Sheet data not found");

            var (startCol, startRow) = ParseCellReference(startCell);
            var (endCol, endRow) = ParseCellReference(endCell);

            // Get or create style index for the formatting
            var styleIndex = GetOrCreateStyleIndex(spreadsheet, formatting);

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var cellRef = GetCellReference(col, row);
                    var cell = GetOrCreateCell(sheetData, cellRef);
                    cell.StyleIndex = styleIndex;
                }
            }

            worksheetPart.Worksheet.Save();
            spreadsheet.Save();
            return new DocumentResult(true, $"Formatting applied to range {startCell}:{endCell}", filePath);
        }
        catch (Exception ex)
        {
            return new DocumentResult(false, $"Failed to format cell range: {ex.Message}");
        }
    }

    public ContentResult GetCellValue(string filePath, string sheetName, string cellReference)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            // Check for encrypted/protected files first
            var protectionCheck = CheckFileProtectionForRead(filePath);
            if (protectionCheck != null)
                return protectionCheck;

            // Use FileStream with sharing to handle OneDrive and other cloud-synced files
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var spreadsheet = OpenSpreadsheet(fileStream, false);
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
        catch (OpenXmlPackageException ex)
        {
            var protectionInfo = OfficeFileProtectionDetector.CheckFileProtection(filePath);
            if (protectionInfo.IsEncrypted || protectionInfo.MayHaveSensitivityLabel)
            {
                return new ContentResult(false, null, OfficeFileProtectionDetector.GetProtectionErrorMessage(protectionInfo, filePath));
            }
            return new ContentResult(false, null, $"Failed to open Excel file (may be corrupted or in an unsupported format): {ex.Message}");
        }
        catch (IOException ex)
        {
            return new ContentResult(false, null, $"Failed to access file (may be locked, cloud-only, or in use): {ex.Message}");
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get cell value: {ex.GetType().Name}: {ex.Message}");
        }
    }

    public ContentResult GetRangeValues(string filePath, string sheetName, string startCell, string endCell)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            // Check for encrypted/protected files first
            var protectionCheck = CheckFileProtectionForRead(filePath);
            if (protectionCheck != null)
                return protectionCheck;

            // Use FileStream with sharing to handle OneDrive and other cloud-synced files
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var spreadsheet = OpenSpreadsheet(fileStream, false);
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
        catch (OpenXmlPackageException ex)
        {
            var protectionInfo = OfficeFileProtectionDetector.CheckFileProtection(filePath);
            if (protectionInfo.IsEncrypted || protectionInfo.MayHaveSensitivityLabel)
            {
                return new ContentResult(false, null, OfficeFileProtectionDetector.GetProtectionErrorMessage(protectionInfo, filePath));
            }
            return new ContentResult(false, null, $"Failed to open Excel file (may be corrupted or in an unsupported format): {ex.Message}");
        }
        catch (IOException ex)
        {
            return new ContentResult(false, null, $"Failed to access file (may be locked, cloud-only, or in use): {ex.Message}");
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get range values: {ex.GetType().Name}: {ex.Message}");
        }
    }

    public ContentResult GetSheetText(string filePath, string sheetName)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            // Check for encrypted/protected files first
            var protectionCheck = CheckFileProtectionForRead(filePath);
            if (protectionCheck != null)
                return protectionCheck;

            // Use FileStream with sharing to handle OneDrive and other cloud-synced files
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var spreadsheet = OpenSpreadsheet(fileStream, false);
            
            return GetSheetTextInternal(spreadsheet, sheetName);
        }
        catch (OpenXmlPackageException ex)
        {
            var protectionInfo = OfficeFileProtectionDetector.CheckFileProtection(filePath);
            if (protectionInfo.IsEncrypted || protectionInfo.MayHaveSensitivityLabel)
            {
                return new ContentResult(false, null, OfficeFileProtectionDetector.GetProtectionErrorMessage(protectionInfo, filePath));
            }
            return new ContentResult(false, null, $"Failed to open Excel file (may be corrupted or in an unsupported format): {ex.Message}");
        }
        catch (IOException ex)
        {
            return new ContentResult(false, null, $"Failed to access file (may be locked, cloud-only, or in use): {ex.Message}");
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get sheet text: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static ContentResult GetSheetTextInternal(SpreadsheetDocument spreadsheet, string sheetName)
    {
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

    public ContentResult GetAllSheetsText(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ContentResult(false, null, $"File not found: {filePath}");

            // Check for encrypted/protected files first
            var protectionCheck = CheckFileProtectionForRead(filePath);
            if (protectionCheck != null)
                return protectionCheck;

            // Use FileStream with sharing to handle OneDrive and other cloud-synced files
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var spreadsheet = OpenSpreadsheet(fileStream, false);
            var workbookPart = spreadsheet.WorkbookPart;
            
            if (workbookPart == null)
                return new ContentResult(false, null, "Workbook part not found - file may not be a valid Excel workbook");

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            if (sheets == null)
                return new ContentResult(false, null, "No sheets found in workbook");

            var sb = new StringBuilder();
            var sheetList = sheets.Elements<Sheet>().ToList();
            
            if (sheetList.Count == 0)
                return new ContentResult(true, "Workbook contains no sheets");

            foreach (var sheet in sheetList)
            {
                sb.AppendLine($"=== Sheet: {sheet.Name} ===");
                // Use internal method to avoid reopening the file
                var result = GetSheetTextInternal(spreadsheet, sheet.Name?.Value ?? string.Empty);
                if (result.Success)
                {
                    sb.AppendLine(result.Content);
                }
                else
                {
                    sb.AppendLine($"Error reading sheet: {result.ErrorMessage}");
                }
                sb.AppendLine();
            }

            return new ContentResult(true, sb.ToString().TrimEnd());
        }
        catch (OpenXmlPackageException ex)
        {
            // If we get here despite the protection check, try to provide more context
            var protectionInfo = OfficeFileProtectionDetector.CheckFileProtection(filePath);
            if (protectionInfo.IsEncrypted || protectionInfo.MayHaveSensitivityLabel)
            {
                return new ContentResult(false, null, OfficeFileProtectionDetector.GetProtectionErrorMessage(protectionInfo, filePath));
            }
            return new ContentResult(false, null, $"Failed to open Excel file (may be corrupted or in an unsupported format): {ex.Message}");
        }
        catch (IOException ex)
        {
            return new ContentResult(false, null, $"Failed to access file (may be locked, cloud-only, or in use): {ex.Message}");
        }
        catch (Exception ex)
        {
            return new ContentResult(false, null, $"Failed to get workbook text: {ex.GetType().Name}: {ex.Message}");
        }
    }

    public DocumentResult DeleteSheet(string filePath, string sheetName)
    {
        try
        {
            using var spreadsheet = OpenSpreadsheet(filePath, true);
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
            using var spreadsheet = OpenSpreadsheet(filePath, true);
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

    public ExcelRangeFormattingResult GetCellFormatting(string filePath, string sheetName, string cellReference)
    {
        return GetRangeFormatting(filePath, sheetName, cellReference, cellReference);
    }

    public ExcelRangeFormattingResult GetRangeFormatting(string filePath, string sheetName, string startCell, string endCell)
    {
        try
        {
            if (!File.Exists(filePath))
                return new ExcelRangeFormattingResult(false, $"File not found: {filePath}", null, null, null);

            // Check for encrypted/protected files first
            var protectionInfo = OfficeFileProtectionDetector.CheckFileProtection(filePath);
            if (protectionInfo.IsEncrypted)
            {
                return new ExcelRangeFormattingResult(false, 
                    OfficeFileProtectionDetector.GetProtectionErrorMessage(protectionInfo, filePath), 
                    null, null, null);
            }

            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var spreadsheet = OpenSpreadsheet(fileStream, false);
            var worksheetPart = GetWorksheetPart(spreadsheet, sheetName);

            if (worksheetPart == null)
                return new ExcelRangeFormattingResult(false, $"Sheet '{sheetName}' not found", null, null, null);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
                return new ExcelRangeFormattingResult(false, "Sheet data not found", null, null, null);

            var (startCol, startRow) = ParseCellReference(startCell);
            var (endCol, endRow) = ParseCellReference(endCell);

            var cells = new List<ExcelCellInfo>();
            var stylesheet = spreadsheet.WorkbookPart?.WorkbookStylesPart?.Stylesheet;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var cellRef = GetCellReference(col, row);
                    var cell = sheetData.Descendants<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRef);

                    string? value = null;
                    string? formula = null;
                    ExcelCellFormattingInfo? formatting = null;

                    if (cell != null)
                    {
                        value = GetCellValueAsString(spreadsheet, cell);
                        formula = cell.CellFormula?.Text;
                        formatting = GetCellFormattingInfo(spreadsheet, cell, stylesheet);
                    }

                    cells.Add(new ExcelCellInfo(cellRef, value, formula, formatting));
                }
            }

            var range = startCell == endCell ? startCell : $"{startCell}:{endCell}";
            return new ExcelRangeFormattingResult(true, null, cells, sheetName, range);
        }
        catch (OpenXmlPackageException ex)
        {
            return new ExcelRangeFormattingResult(false, $"Failed to open Excel file: {ex.Message}", null, null, null);
        }
        catch (Exception ex)
        {
            return new ExcelRangeFormattingResult(false, $"Failed to get formatting: {ex.Message}", null, null, null);
        }
    }

    private static ExcelCellFormattingInfo? GetCellFormattingInfo(SpreadsheetDocument document, Cell cell, Stylesheet? stylesheet)
    {
        if (stylesheet == null || cell.StyleIndex == null)
            return null;

        var styleIndex = (int)cell.StyleIndex.Value;
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null || styleIndex >= cellFormats.Count?.Value)
            return null;

        var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault(styleIndex);
        if (cellFormat == null)
            return null;

        // Get font info
        bool bold = false;
        bool italic = false;
        bool underline = false;
        string? fontName = null;
        double? fontSize = null;
        string? fontColor = null;

        if (cellFormat.FontId != null)
        {
            var fonts = stylesheet.Fonts;
            var font = fonts?.Elements<Font>().ElementAtOrDefault((int)cellFormat.FontId.Value);
            if (font != null)
            {
                bold = font.Bold != null;
                italic = font.Italic != null;
                underline = font.Underline != null;
                fontName = font.FontName?.Val?.Value;
                fontSize = font.FontSize?.Val?.Value;
                
                var color = font.Color;
                if (color?.Rgb != null)
                {
                    fontColor = "#" + color.Rgb.Value?[2..]; // Remove alpha prefix
                }
            }
        }

        // Get fill info
        string? backgroundColor = null;
        if (cellFormat.FillId != null && cellFormat.FillId.Value > 1) // 0 and 1 are reserved
        {
            var fills = stylesheet.Fills;
            var fill = fills?.Elements<Fill>().ElementAtOrDefault((int)cellFormat.FillId.Value);
            var patternFill = fill?.PatternFill;
            if (patternFill?.ForegroundColor?.Rgb != null)
            {
                var rgb = patternFill.ForegroundColor.Rgb.Value;
                if (rgb != null && rgb.Length >= 6)
                {
                    backgroundColor = "#" + (rgb.Length == 8 ? rgb[2..] : rgb);
                }
            }
        }

        // Get border info
        bool hasBorder = false;
        string? borderStyle = null;
        if (cellFormat.BorderId != null && cellFormat.BorderId.Value > 0)
        {
            var borders = stylesheet.Borders;
            var border = borders?.Elements<Border>().ElementAtOrDefault((int)cellFormat.BorderId.Value);
            if (border != null)
            {
                var leftStyle = border.LeftBorder?.Style?.Value;
                if (leftStyle != null && leftStyle != BorderStyleValues.None)
                {
                    hasBorder = true;
                    borderStyle = leftStyle.ToString();
                }
            }
        }

        // Get alignment info
        string? horizontalAlignment = null;
        string? verticalAlignment = null;
        bool wrapText = false;

        var alignment = cellFormat.Alignment;
        if (alignment != null)
        {
            if (alignment.Horizontal != null && alignment.Horizontal.HasValue)
                horizontalAlignment = alignment.Horizontal.Value.ToString();
            if (alignment.Vertical != null && alignment.Vertical.HasValue)
                verticalAlignment = alignment.Vertical.Value.ToString();
            wrapText = alignment.WrapText?.Value ?? false;
        }

        // Get number format
        string? numberFormat = null;
        if (cellFormat.NumberFormatId != null)
        {
            var numFormatId = cellFormat.NumberFormatId.Value;
            // Check built-in formats
            numberFormat = numFormatId switch
            {
                0 => "General",
                1 => "0",
                2 => "0.00",
                3 => "#,##0",
                4 => "#,##0.00",
                9 => "0%",
                10 => "0.00%",
                11 => "0.00E+00",
                14 => "mm-dd-yy",
                15 => "d-mmm-yy",
                16 => "d-mmm",
                17 => "mmm-yy",
                18 => "h:mm AM/PM",
                19 => "h:mm:ss AM/PM",
                20 => "h:mm",
                21 => "h:mm:ss",
                22 => "m/d/yy h:mm",
                _ => null
            };

            // Check custom formats
            if (numberFormat == null && numFormatId >= 164)
            {
                var numFormats = stylesheet.NumberingFormats;
                var customFormat = numFormats?.Elements<NumberingFormat>()
                    .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFormatId);
                numberFormat = customFormat?.FormatCode?.Value;
            }
        }

        return new ExcelCellFormattingInfo(
            Bold: bold,
            Italic: italic,
            Underline: underline,
            FontName: fontName,
            FontSize: fontSize,
            FontColor: fontColor,
            BackgroundColor: backgroundColor,
            NumberFormat: numberFormat,
            HorizontalAlignment: horizontalAlignment ?? "General",
            VerticalAlignment: verticalAlignment ?? "Bottom",
            WrapText: wrapText,
            HasBorder: hasBorder,
            BorderStyle: borderStyle
        );
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
        var styleIndex = GetOrCreateStyleIndex(document, formatting);
        cell.StyleIndex = styleIndex;
    }

    private static uint GetOrCreateStyleIndex(SpreadsheetDocument document, ExcelCellFormatting formatting)
    {
        var workbookPart = document.WorkbookPart;
        var stylesPart = workbookPart?.WorkbookStylesPart;
        
        if (stylesPart?.Stylesheet == null)
            return 0;

        var stylesheet = stylesPart.Stylesheet;

        // Get or create font
        uint fontIndex = 0;
        if (formatting.Bold || formatting.Italic || !string.IsNullOrEmpty(formatting.FontColor))
        {
            var fonts = stylesheet.Fonts;
            if (fonts == null)
            {
                fonts = new Fonts();
                stylesheet.InsertAt(fonts, 0);
            }
            
            var newFont = new Font();
            
            if (formatting.Bold)
                newFont.Append(new Bold());
            if (formatting.Italic)
                newFont.Append(new Italic());
            if (!string.IsNullOrEmpty(formatting.FontColor))
            {
                newFont.Append(new Color { Rgb = new HexBinaryValue(NormalizeColor(formatting.FontColor)) });
            }
            newFont.Append(new FontSize { Val = 11 });
            newFont.Append(new FontName { Val = "Calibri" });

            // Get current count BEFORE appending
            var currentFontCount = (uint)fonts.Elements<Font>().Count();
            fonts.Append(newFont);
            fontIndex = currentFontCount; // New font is at this index
            fonts.Count = currentFontCount + 1;
        }

        // Get or create fill
        uint fillIndex = 0;
        if (!string.IsNullOrEmpty(formatting.BackgroundColor))
        {
            var fills = stylesheet.Fills;
            if (fills == null)
            {
                // Create fills with required first two entries
                fills = new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
                );
                fills.Count = 2;
                stylesheet.InsertAfter(fills, stylesheet.Fonts);
            }
            
            var patternFill = new PatternFill { PatternType = PatternValues.Solid };
            patternFill.ForegroundColor = new ForegroundColor { Rgb = new HexBinaryValue(NormalizeColor(formatting.BackgroundColor)) };
            patternFill.BackgroundColor = new BackgroundColor { Indexed = 64 };
            
            var newFill = new Fill(patternFill);
            
            // Get current count BEFORE appending
            var currentFillCount = (uint)fills.Elements<Fill>().Count();
            fills.Append(newFill);
            fillIndex = currentFillCount; // New fill is at this index
            fills.Count = currentFillCount + 1;
        }

        // Get or create border
        uint borderIndex = 0;
        if (!string.IsNullOrEmpty(formatting.BorderStyle))
        {
            var borders = stylesheet.Borders;
            if (borders == null)
            {
                borders = new Borders(
                    new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()
                    )
                );
                borders.Count = 1;
                stylesheet.InsertAfter(borders, stylesheet.Fills);
            }
            
            var borderStyle = formatting.BorderStyle.ToLowerInvariant() switch
            {
                "thin" => BorderStyleValues.Thin,
                "medium" => BorderStyleValues.Medium,
                "thick" => BorderStyleValues.Thick,
                "double" => BorderStyleValues.Double,
                "dashed" => BorderStyleValues.Dashed,
                "dotted" => BorderStyleValues.Dotted,
                _ => BorderStyleValues.Thin
            };

            var newBorder = new Border(
                new LeftBorder { Style = borderStyle },
                new RightBorder { Style = borderStyle },
                new TopBorder { Style = borderStyle },
                new BottomBorder { Style = borderStyle },
                new DiagonalBorder()
            );
            
            // Get current count BEFORE appending
            var currentBorderCount = (uint)borders.Elements<Border>().Count();
            borders.Append(newBorder);
            borderIndex = currentBorderCount; // New border is at this index
            borders.Count = currentBorderCount + 1;
        }

        // Create cell format
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null)
        {
            cellFormats = new CellFormats(
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }
            );
            cellFormats.Count = 1;
            stylesheet.Append(cellFormats);
        }
        
        var cellFormat = new CellFormat
        {
            FontId = fontIndex,
            FillId = fillIndex,
            BorderId = borderIndex,
            ApplyFont = formatting.Bold || formatting.Italic || !string.IsNullOrEmpty(formatting.FontColor),
            ApplyFill = !string.IsNullOrEmpty(formatting.BackgroundColor),
            ApplyBorder = !string.IsNullOrEmpty(formatting.BorderStyle)
        };

        // Apply alignment
        if (formatting.HorizontalAlignment != "General" || formatting.VerticalAlignment != "Bottom" || formatting.WrapText)
        {
            var alignment = new Alignment
            {
                WrapText = formatting.WrapText
            };

            alignment.Horizontal = formatting.HorizontalAlignment.ToLowerInvariant() switch
            {
                "left" => HorizontalAlignmentValues.Left,
                "center" => HorizontalAlignmentValues.Center,
                "right" => HorizontalAlignmentValues.Right,
                "justify" => HorizontalAlignmentValues.Justify,
                _ => null
            };

            alignment.Vertical = formatting.VerticalAlignment.ToLowerInvariant() switch
            {
                "top" => VerticalAlignmentValues.Top,
                "center" => VerticalAlignmentValues.Center,
                "bottom" => VerticalAlignmentValues.Bottom,
                _ => null
            };

            cellFormat.Append(alignment);
            cellFormat.ApplyAlignment = true;
        }

        // Apply number format if specified
        if (!string.IsNullOrEmpty(formatting.NumberFormat))
        {
            // Add custom number format
            var numFormats = stylesheet.NumberingFormats;
            if (numFormats == null)
            {
                numFormats = new NumberingFormats();
                stylesheet.InsertAt(numFormats, 0);
                numFormats.Count = 0;
            }

            uint formatId = 164 + (uint)(numFormats.Count?.Value ?? 0); // Custom formats start at 164
            numFormats.Append(new NumberingFormat { NumberFormatId = formatId, FormatCode = formatting.NumberFormat });
            numFormats.Count = (numFormats.Count?.Value ?? 0) + 1;

            cellFormat.NumberFormatId = formatId;
            cellFormat.ApplyNumberFormat = true;
        }

        // Get current count BEFORE appending
        var currentCellFormatCount = (uint)cellFormats.Elements<CellFormat>().Count();
        cellFormats.Append(cellFormat);
        var styleIndex = currentCellFormatCount; // New cell format is at this index
        cellFormats.Count = currentCellFormatCount + 1;

        stylesheet.Save();
        return styleIndex;
    }

    private static string NormalizeColor(string color)
    {
        // Ensure color is in ARGB format (8 hex chars)
        if (string.IsNullOrEmpty(color))
            return "FF000000";
        
        color = color.TrimStart('#');
        
        // If RGB (6 chars), add FF for alpha
        if (color.Length == 6)
            return "FF" + color.ToUpperInvariant();
        
        // If already ARGB (8 chars), return as-is
        if (color.Length == 8)
            return color.ToUpperInvariant();
        
        return "FF000000"; // Default to black if invalid
    }

    private static void ExpandTablesForRange(WorksheetPart worksheetPart, int startCol, int startRow, int endCol, int endRow)
    {
        foreach (var tableDefPart in worksheetPart.TableDefinitionParts)
        {
            var table = tableDefPart.Table;
            if (table?.Reference?.Value == null) continue;

            var tableRange = table.Reference.Value;
            var parts = tableRange.Split(':');
            if (parts.Length != 2) continue;

            var (tableStartCol, tableStartRow) = ParseCellReference(parts[0]);
            var (tableEndCol, tableEndRow) = ParseCellReference(parts[1]);

            // Check if the new data range overlaps with or is adjacent to the table
            // The table should be expanded if:
            // 1. New data starts within the table's column range
            // 2. New data is in the row immediately after the table's last row (appending)
            // 3. New data overlaps with existing table range
            
            bool columnsOverlap = startCol <= tableEndCol && endCol >= tableStartCol;
            bool rowsOverlapOrAdjacent = startRow <= tableEndRow + 1 && endRow >= tableStartRow;
            
            if (columnsOverlap && rowsOverlapOrAdjacent)
            {
                // Expand table to include new range
                var newStartCol = Math.Min(tableStartCol, startCol);
                var newStartRow = Math.Min(tableStartRow, startRow);
                var newEndCol = Math.Max(tableEndCol, endCol);
                var newEndRow = Math.Max(tableEndRow, endRow);

                var newStartCellRef = GetCellReference(newStartCol, newStartRow);
                var newEndCellRef = GetCellReference(newEndCol, newEndRow);
                var newRange = $"{newStartCellRef}:{newEndCellRef}";

                table.Reference = newRange;

                // Update AutoFilter if present
                var autoFilter = table.AutoFilter;
                if (autoFilter != null)
                {
                    autoFilter.Reference = newRange;
                }

                // Update table columns if columns were added
                var tableColumns = table.TableColumns;
                if (tableColumns != null)
                {
                    var currentColCount = (int)(tableColumns.Count?.Value ?? 0);
                    var newColCount = newEndCol - newStartCol + 1;

                    for (int i = currentColCount; i < newColCount; i++)
                    {
                        tableColumns.Append(new TableColumn { Id = (uint)(i + 1), Name = $"Column{i + 1}" });
                    }
                    tableColumns.Count = (uint)newColCount;
                }

                tableDefPart.Table.Save();
            }
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
