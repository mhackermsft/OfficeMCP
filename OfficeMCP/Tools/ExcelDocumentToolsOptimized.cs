using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// AI-Optimized MCP Tools for Excel workbooks. Reduces tool calls through batch operations.
/// </summary>
[McpServerToolType]
public sealed class ExcelDocumentToolsOptimized(IExcelDocumentService excelService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    [McpServerTool, Description(@"Create a new Excel workbook and optionally add initial data in a single call.

**Examples**:
- Simple: {""filePath"": ""C:/data/sales.xlsx""}
- With sheet name: {""filePath"": ""C:/data/sales.xlsx"", ""sheetName"": ""Q1 Sales""}
- With initial data: {""filePath"": ""C:/data/sales.xlsx"", ""initialData"": [[""Product"",""Sales""],[""Widget"",""100""]], ""hasHeaders"": true}")]
    public string CreateExcelWorkbook(
        [Description("Full file path for the new workbook (e.g., C:/Documents/data.xlsx)")] string filePath,
        [Description("Name for the first sheet (default: 'Sheet1')")] string? sheetName = null,
        [Description("Initial data as JSON 2D array (e.g., [[\"Header1\",\"Header2\"],[\"Row1Col1\",\"Row1Col2\"]])")] string? initialDataJson = null,
        [Description("If true, first row of initialData is treated as headers")] bool hasHeaders = true)
    {
        var result = excelService.CreateWorkbook(filePath, sheetName);

        if (!result.Success)
        {
            return JsonSerializer.Serialize(result, JsonOptions);
        }

        // If initial data is provided, add it as a table
        if (!string.IsNullOrWhiteSpace(initialDataJson))
        {
            try
            {
                var data = JsonSerializer.Deserialize<string[][]>(initialDataJson, JsonOptions);
                if (data != null && data.Length > 0)
                {
                    var tableResult = excelService.AddTable(filePath, sheetName ?? "Sheet1", "A1", data, hasHeaders);
                    if (!tableResult.Success)
                    {
                        return JsonSerializer.Serialize(new DocumentResult(false, $"Workbook created but data failed: {tableResult.Message}", filePath), JsonOptions);
                    }
                    return JsonSerializer.Serialize(new DocumentResult(true, "Workbook created with initial data", filePath), JsonOptions);
                }
            }
            catch (JsonException ex)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, $"Workbook created but invalid data JSON: {ex.Message}", filePath), JsonOptions);
            }
        }

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Set cell values in an Excel workbook. Use this for simple data entry without JSON escaping.")]
    public string SetExcelCells(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Starting cell reference (e.g., 'A1')")] string startCell,
        [Description("Data as JSON 2D array (e.g., [[\"Name\",\"Value\"],[\"Item1\",\"100\"]])")] string dataJson,
        [Description("If true, format as a table with headers")] bool asTable = false,
        [Description("If asTable is true, first row contains headers")] bool hasHeaders = true)
    {
        try
        {
            var data = JsonSerializer.Deserialize<string[][]>(dataJson, JsonOptions);
            if (data == null || data.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty data array"), JsonOptions);
            }

            DocumentResult result;
            if (asTable)
            {
                result = excelService.AddTable(filePath, sheetName, startCell, data, hasHeaders);
            }
            else
            {
                result = excelService.SetRangeValues(filePath, sheetName, startCell, data);
            }
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description(@"Add a formula to a cell in an Excel workbook.")]
    public string AddExcelFormula(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Cell reference for the formula (e.g., 'C1')")] string cellReference,
        [Description("Excel formula without the leading '=' (e.g., 'SUM(A1:B1)', 'AVERAGE(A1:A10)')")] string formula)
    {
        var result = excelService.AddFormula(filePath, sheetName, cellReference, formula);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Manage sheets in an Excel workbook (add, delete, rename).")]
    public string ManageExcelSheet(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Operation: 'add', 'delete', or 'rename'")] string operation,
        [Description("Sheet name (for add/delete) or current name (for rename)")] string sheetName,
        [Description("New name for the sheet (required for 'rename' operation)")] string? newName = null)
    {
        DocumentResult result = operation.ToLowerInvariant() switch
        {
            "add" => excelService.AddSheet(filePath, sheetName),
            "delete" => excelService.DeleteSheet(filePath, sheetName),
            "rename" when !string.IsNullOrEmpty(newName) => excelService.RenameSheet(filePath, sheetName, newName),
            "rename" => new DocumentResult(false, "New name is required for rename operation"),
            _ => new DocumentResult(false, $"Unknown operation: {operation}. Use 'add', 'delete', or 'rename'.")
        };
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description(@"Perform batch operations on an Excel workbook using a JSON array.

**Operations JSON array** - each object has a 'type' and type-specific properties:

| Type | Key Properties |
|------|----------------|
| addSheet | sheetName |
| deleteSheet | sheetName |
| renameSheet | sheetName, newSheetName |
| setCellValue | sheetName, cellReference, value, bold, italic |
| setRangeValues | sheetName, startCell, values (2D array) |
| addTable | sheetName, startCell, tableData (2D array), hasHeaders |
| addFormula | sheetName, cellReference, formula (without =) |
| mergeCells | sheetName, startCell, endCell |
| setColumnWidth | sheetName, columnIndex (1=A), width |
| setRowHeight | sheetName, rowIndex (1-based), height |
| addImage | sheetName, imagePath, cellReference, widthInches, heightInches |")]
    public string BatchModifyExcelWorkbook(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("JSON array of operations")] string operationsJson)
    {
        try
        {
            var operations = JsonSerializer.Deserialize<ExcelOperation[]>(operationsJson, JsonOptions);
            if (operations == null || operations.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "No operations provided"), JsonOptions);
            }

            var details = new List<OperationOutcome>();
            int successCount = 0;
            int failCount = 0;

            for (int i = 0; i < operations.Length; i++)
            {
                var op = operations[i];
                DocumentResult? opResult = null;

                try
                {
                    opResult = op.Type.ToLowerInvariant() switch
                    {
                        "addsheet" => ProcessAddSheet(filePath, op),
                        "deletesheet" => ProcessDeleteSheet(filePath, op),
                        "renamesheet" => ProcessRenameSheet(filePath, op),
                        "setcellvalue" => ProcessSetCellValue(filePath, op),
                        "setrangevalues" => ProcessSetRangeValues(filePath, op),
                        "addtable" => ProcessAddTable(filePath, op),
                        "addformula" => ProcessAddFormula(filePath, op),
                        "mergecells" => ProcessMergeCells(filePath, op),
                        "setcolumnwidth" => ProcessSetColumnWidth(filePath, op),
                        "setrowheight" => ProcessSetRowHeight(filePath, op),
                        "addimage" => ProcessAddImage(filePath, op),
                        _ => new DocumentResult(false, $"Unknown operation type: {op.Type}")
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

            var batchResult = new BatchOperationResult(
                Success: failCount == 0,
                Message: failCount == 0
                    ? $"All {successCount} operations completed successfully"
                    : $"{successCount} succeeded, {failCount} failed",
                TotalOperations: operations.Length,
                SuccessfulOperations: successCount,
                FailedOperations: failCount,
                Details: details
            );

            return JsonSerializer.Serialize(batchResult, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description(@"Read content from an Excel workbook.

**Options for 'readType' parameter**:
- 'allSheets' (default): Get all content from all sheets
- 'sheet': Get all content from a specific sheet (use sheetName)
- 'cell': Get single cell value (use sheetName, cellReference)
- 'range': Get range of cells (use sheetName, startCell, endCell)

**Examples**:
- Get all: {""filePath"": ""C:/data/sales.xlsx""}
- Get sheet: {""filePath"": ""C:/data/sales.xlsx"", ""readType"": ""sheet"", ""sheetName"": ""Sales""}
- Get cell: {""filePath"": ""C:/data/sales.xlsx"", ""readType"": ""cell"", ""sheetName"": ""Sales"", ""cellReference"": ""A1""}
- Get range: {""filePath"": ""C:/data/sales.xlsx"", ""readType"": ""range"", ""sheetName"": ""Sales"", ""startCell"": ""A1"", ""endCell"": ""D10""}")]
    public string ReadExcelWorkbook(
        [Description("Path to the Excel workbook")] string filePath,
        [Description("Type of read: 'allSheets', 'sheet', 'cell', or 'range'")] string readType = "allSheets",
        [Description("Sheet name (required for sheet, cell, range read types)")] string? sheetName = null,
        [Description("Cell reference for 'cell' read type (e.g., 'A1')")] string? cellReference = null,
        [Description("Start cell for 'range' read type")] string? startCell = null,
        [Description("End cell for 'range' read type")] string? endCell = null)
    {
        ContentResult result = readType.ToLowerInvariant() switch
        {
            "sheet" when !string.IsNullOrEmpty(sheetName) => excelService.GetSheetText(filePath, sheetName),
            "cell" when !string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(cellReference) =>
                excelService.GetCellValue(filePath, sheetName, cellReference),
            "range" when !string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(startCell) && !string.IsNullOrEmpty(endCell) =>
                excelService.GetRangeValues(filePath, sheetName, startCell, endCell),
            _ => excelService.GetAllSheetsText(filePath)
        };

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    #region Private Operation Processors

    private DocumentResult ProcessAddSheet(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName))
        {
            return new DocumentResult(false, "Sheet name is required");
        }
        return excelService.AddSheet(filePath, op.SheetName);
    }

    private DocumentResult ProcessDeleteSheet(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName))
        {
            return new DocumentResult(false, "Sheet name is required");
        }
        return excelService.DeleteSheet(filePath, op.SheetName);
    }

    private DocumentResult ProcessRenameSheet(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.NewSheetName))
        {
            return new DocumentResult(false, "Sheet name and new sheet name are required");
        }
        return excelService.RenameSheet(filePath, op.SheetName, op.NewSheetName);
    }

    private DocumentResult ProcessSetCellValue(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.CellReference))
        {
            return new DocumentResult(false, "Sheet name and cell reference are required");
        }
        var formatting = new ExcelCellFormatting(
            Bold: op.Bold ?? false,
            Italic: op.Italic ?? false,
            WrapText: op.WrapText ?? false
        );
        return excelService.SetCellValue(filePath, op.SheetName, op.CellReference, op.Value ?? "", formatting);
    }

    private DocumentResult ProcessSetRangeValues(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.StartCell) || op.Values == null)
        {
            return new DocumentResult(false, "Sheet name, start cell, and values are required");
        }
        return excelService.SetRangeValues(filePath, op.SheetName, op.StartCell, op.Values);
    }

    private DocumentResult ProcessAddTable(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.StartCell) || op.TableData == null)
        {
            return new DocumentResult(false, "Sheet name, start cell, and table data are required");
        }
        return excelService.AddTable(filePath, op.SheetName, op.StartCell, op.TableData, op.HasHeaders ?? true);
    }

    private DocumentResult ProcessAddFormula(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.CellReference) || string.IsNullOrWhiteSpace(op.Formula))
        {
            return new DocumentResult(false, "Sheet name, cell reference, and formula are required");
        }
        return excelService.AddFormula(filePath, op.SheetName, op.CellReference, op.Formula);
    }

    private DocumentResult ProcessMergeCells(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.StartCell) || string.IsNullOrWhiteSpace(op.EndCell))
        {
            return new DocumentResult(false, "Sheet name, start cell, and end cell are required");
        }
        return excelService.MergeCells(filePath, op.SheetName, op.StartCell, op.EndCell);
    }

    private DocumentResult ProcessSetColumnWidth(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || !op.ColumnIndex.HasValue || !op.Width.HasValue)
        {
            return new DocumentResult(false, "Sheet name, column index, and width are required");
        }
        return excelService.SetColumnWidth(filePath, op.SheetName, op.ColumnIndex.Value, op.Width.Value);
    }

    private DocumentResult ProcessSetRowHeight(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || !op.RowIndex.HasValue || !op.Height.HasValue)
        {
            return new DocumentResult(false, "Sheet name, row index, and height are required");
        }
        return excelService.SetRowHeight(filePath, op.SheetName, op.RowIndex.Value, op.Height.Value);
    }

    private DocumentResult ProcessAddImage(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.ImagePath) || string.IsNullOrWhiteSpace(op.CellReference))
        {
            return new DocumentResult(false, "Sheet name, image path, and cell reference are required");
        }
        var options = new ImageOptions(
            WidthEmu: (long)((op.WidthInches ?? 4.0) * 914400),
            HeightEmu: (long)((op.HeightInches ?? 3.0) * 914400),
            AltText: op.AltText
        );
        return excelService.AddImage(filePath, op.SheetName, op.ImagePath, op.CellReference, options);
    }

    #endregion
}
