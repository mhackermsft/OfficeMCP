using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// AI-Optimized MCP Tools for Excel workbooks.
/// Consolidated tools with simplified descriptions and tool annotations for better AI discoverability.
/// </summary>
[McpServerToolType]
public sealed class ExcelDocumentToolsOptimized(IExcelDocumentService excelService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    #region Core Workbook Operations

    [McpServerTool(Name = "excel_create", Destructive = false, ReadOnly = false), Description("Creates an Excel workbook (.xlsx) with optional initial data as a table.")]
    public string CreateExcelWorkbook(
        [Description("Full path (e.g., C:/data/sales.xlsx)")] string filePath,
        [Description("Sheet name (default: Sheet1)")] string? sheetName = null,
        [Description("Initial data as JSON 2D array: [[\"Header1\",\"Header2\"],[\"Val1\",\"Val2\"]]")] string? initialDataJson = null,
        [Description("Treat first row as headers")] bool hasHeaders = true)
    {
        var result = excelService.CreateWorkbook(filePath, sheetName);
        if (!result.Success)
            return JsonSerializer.Serialize(result, JsonOptions);

        if (!string.IsNullOrWhiteSpace(initialDataJson))
        {
            try
            {
                var data = JsonSerializer.Deserialize<string[][]>(initialDataJson, JsonOptions);
                if (data != null && data.Length > 0)
                {
                    var tableResult = excelService.AddTable(filePath, sheetName ?? "Sheet1", "A1", data, hasHeaders);
                    if (!tableResult.Success)
                        return JsonSerializer.Serialize(new DocumentResult(false, $"Created but data failed: {tableResult.Message}", filePath), JsonOptions);
                    return JsonSerializer.Serialize(new DocumentResult(true, "Workbook created with data", filePath), JsonOptions);
                }
            }
            catch (JsonException ex)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, $"Created but invalid JSON: {ex.Message}", filePath), JsonOptions);
            }
        }

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "excel_read", Destructive = false, ReadOnly = true), Description("Reads content from an Excel workbook. Returns all sheets by default, or specific sheet/cell/range.")]
    public string ReadExcelWorkbook(
        [Description("Path to the workbook")] string filePath,
        [Description("allSheets (default), sheet, cell, or range")] string readType = "allSheets",
        [Description("Sheet name (required for sheet/cell/range)")] string? sheetName = null,
        [Description("Cell reference for 'cell' mode (e.g., A1)")] string? cellReference = null,
        [Description("Start cell for 'range' mode")] string? startCell = null,
        [Description("End cell for 'range' mode")] string? endCell = null)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new ContentResult(false, null, $"File not found: {filePath}. Use excel_create to create a workbook first."), JsonOptions);

        ContentResult result = readType.ToLowerInvariant() switch
        {
            "sheet" when !string.IsNullOrEmpty(sheetName) => excelService.GetSheetText(filePath, sheetName),
            "sheet" => new ContentResult(false, null, "sheetName is required when readType is 'sheet'"),
            "cell" when !string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(cellReference) =>
                excelService.GetCellValue(filePath, sheetName, cellReference),
            "cell" => new ContentResult(false, null, "sheetName and cellReference are required when readType is 'cell'"),
            "range" when !string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(startCell) && !string.IsNullOrEmpty(endCell) =>
                excelService.GetRangeValues(filePath, sheetName, startCell, endCell),
            "range" => new ContentResult(false, null, "sheetName, startCell, and endCell are required when readType is 'range'"),
            _ => excelService.GetAllSheetsText(filePath)
        };

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "excel_set_cells", Destructive = false, ReadOnly = false), Description("Sets cell values in an Excel workbook. Can format as table.")]
    public string SetExcelCells(
        [Description("Path to the workbook")] string filePath,
        [Description("Sheet name")] string sheetName,
        [Description("Starting cell (e.g., A1)")] string startCell,
        [Description("Data as JSON 2D array: [[\"Name\",\"Value\"],[\"Item1\",\"100\"]]")] string dataJson,
        [Description("Format as table with headers")] bool asTable = false,
        [Description("First row contains headers")] bool hasHeaders = true)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {filePath}", Suggestion: "Use excel_create to create the workbook first"), JsonOptions);

        try
        {
            var data = JsonSerializer.Deserialize<string[][]>(dataJson, JsonOptions);
            if (data == null || data.Length == 0)
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty data array", Suggestion: "Provide data as JSON 2D array: [[\"A\",\"B\"],[\"1\",\"2\"]]"), JsonOptions);

            var result = asTable
                ? excelService.AddTable(filePath, sheetName, startCell, data, hasHeaders)
                : excelService.SetRangeValues(filePath, sheetName, startCell, data);
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON: {ex.Message}", Suggestion: "Ensure dataJson is a valid 2D JSON array"), JsonOptions);
        }
    }

    #endregion

    #region Sheet and Cell Operations

    [McpServerTool(Name = "excel_formula", Destructive = false, ReadOnly = false), Description("Adds a formula to a cell. Formula without leading '=' (e.g., SUM(A1:B1)).")]
    public string AddExcelFormula(
        [Description("Path to the workbook")] string filePath,
        [Description("Sheet name")] string sheetName,
        [Description("Cell reference (e.g., C1)")] string cellReference,
        [Description("Formula without '=' (e.g., SUM(A1:B1))")] string formula)
    {
        var result = excelService.AddFormula(filePath, sheetName, cellReference, formula);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool(Name = "excel_manage_sheet", Destructive = true, ReadOnly = false), Description("Manages sheets: add, delete, or rename.")]
    public string ManageExcelSheet(
        [Description("Path to the workbook")] string filePath,
        [Description("add, delete, or rename")] string operation,
        [Description("Sheet name (or current name for rename)")] string sheetName,
        [Description("New name (required for rename)")] string? newName = null)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {filePath}", Suggestion: "Use excel_create to create the workbook first"), JsonOptions);

        DocumentResult result = operation.ToLowerInvariant() switch
        {
            "add" => excelService.AddSheet(filePath, sheetName),
            "delete" => excelService.DeleteSheet(filePath, sheetName),
            "rename" when !string.IsNullOrEmpty(newName) => excelService.RenameSheet(filePath, sheetName, newName),
            "rename" => new DocumentResult(false, "New name is required for rename", Suggestion: "Provide the 'newName' parameter"),
            _ => new DocumentResult(false, $"Unknown operation: '{operation}'", Suggestion: "Valid operations: add, delete, rename")
        };
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    #endregion

    #region Batch Operations

    [McpServerTool(Name = "excel_batch", Destructive = true, ReadOnly = false), Description("Performs multiple operations. Types: addSheet, deleteSheet, renameSheet, setCellValue, setRangeValues, addTable, addFormula, mergeCells, setColumnWidth, setRowHeight, addImage.")]
    public string BatchModifyExcelWorkbook(
        [Description("Path to the workbook")] string filePath,
        [Description("JSON array: [{\"type\":\"addSheet\",\"sheetName\":\"Data\"}, {\"type\":\"addFormula\",\"sheetName\":\"Data\",\"cellReference\":\"C1\",\"formula\":\"SUM(A1:B1)\"}]")] string operationsJson)
    {
        if (!File.Exists(filePath))
            return JsonSerializer.Serialize(new DocumentResult(false, $"File not found: {filePath}", Suggestion: "Use excel_create to create the workbook first"), JsonOptions);

        try
        {
            var operations = JsonSerializer.Deserialize<ExcelOperation[]>(operationsJson, JsonOptions);
            if (operations == null || operations.Length == 0)
                return JsonSerializer.Serialize(new DocumentResult(false, "No operations provided", Suggestion: "Provide a JSON array of operations"), JsonOptions);

            var details = new List<OperationOutcome>();
            int successCount = 0, failCount = 0;

            for (int i = 0; i < operations.Length; i++)
            {
                var op = operations[i];
                try
                {
                    var opResult = op.Type.ToLowerInvariant() switch
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
                        _ => new DocumentResult(false, $"Unknown type: '{op.Type}'", Suggestion: "Valid types: addSheet, deleteSheet, renameSheet, setCellValue, setRangeValues, addTable, addFormula, mergeCells, setColumnWidth, setRowHeight, addImage")
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

            return JsonSerializer.Serialize(new BatchOperationResult(
                Success: failCount == 0,
                Message: failCount == 0 ? $"All {successCount} operations completed" : $"{successCount} succeeded, {failCount} failed",
                TotalOperations: operations.Length,
                SuccessfulOperations: successCount,
                FailedOperations: failCount,
                Details: details
            ), JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON: {ex.Message}", Suggestion: "Ensure operationsJson is a valid JSON array"), JsonOptions);
        }
    }

    #endregion

    #region Private Operation Processors

    private DocumentResult ProcessAddSheet(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName))
            return new DocumentResult(false, "Sheet name is required");
        return excelService.AddSheet(filePath, op.SheetName);
    }

    private DocumentResult ProcessDeleteSheet(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName))
            return new DocumentResult(false, "Sheet name is required");
        return excelService.DeleteSheet(filePath, op.SheetName);
    }

    private DocumentResult ProcessRenameSheet(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.NewSheetName))
            return new DocumentResult(false, "Sheet name and new name are required");
        return excelService.RenameSheet(filePath, op.SheetName, op.NewSheetName);
    }

    private DocumentResult ProcessSetCellValue(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.CellReference))
            return new DocumentResult(false, "Sheet name and cell reference are required");
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
            return new DocumentResult(false, "Sheet name, start cell, and values are required");
        return excelService.SetRangeValues(filePath, op.SheetName, op.StartCell, op.Values);
    }

    private DocumentResult ProcessAddTable(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.StartCell) || op.TableData == null)
            return new DocumentResult(false, "Sheet name, start cell, and table data are required");
        return excelService.AddTable(filePath, op.SheetName, op.StartCell, op.TableData, op.HasHeaders ?? true);
    }

    private DocumentResult ProcessAddFormula(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.CellReference) || string.IsNullOrWhiteSpace(op.Formula))
            return new DocumentResult(false, "Sheet name, cell reference, and formula are required");
        return excelService.AddFormula(filePath, op.SheetName, op.CellReference, op.Formula);
    }

    private DocumentResult ProcessMergeCells(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.StartCell) || string.IsNullOrWhiteSpace(op.EndCell))
            return new DocumentResult(false, "Sheet name, start cell, and end cell are required");
        return excelService.MergeCells(filePath, op.SheetName, op.StartCell, op.EndCell);
    }

    private DocumentResult ProcessSetColumnWidth(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || !op.ColumnIndex.HasValue || !op.Width.HasValue)
            return new DocumentResult(false, "Sheet name, column index, and width are required");
        return excelService.SetColumnWidth(filePath, op.SheetName, op.ColumnIndex.Value, op.Width.Value);
    }

    private DocumentResult ProcessSetRowHeight(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || !op.RowIndex.HasValue || !op.Height.HasValue)
            return new DocumentResult(false, "Sheet name, row index, and height are required");
        return excelService.SetRowHeight(filePath, op.SheetName, op.RowIndex.Value, op.Height.Value);
    }

    private DocumentResult ProcessAddImage(string filePath, ExcelOperation op)
    {
        if (string.IsNullOrWhiteSpace(op.SheetName) || string.IsNullOrWhiteSpace(op.ImagePath) || string.IsNullOrWhiteSpace(op.CellReference))
            return new DocumentResult(false, "Sheet name, image path, and cell reference are required");
        var options = new ImageOptions(
            WidthEmu: (long)((op.WidthInches ?? 4.0) * 914400),
            HeightEmu: (long)((op.HeightInches ?? 3.0) * 914400),
            AltText: op.AltText
        );
        return excelService.AddImage(filePath, op.SheetName, op.ImagePath, op.CellReference, options);
    }

    #endregion
}
