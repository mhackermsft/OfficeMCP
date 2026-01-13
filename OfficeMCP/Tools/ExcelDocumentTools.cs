using ModelContextProtocol.Server;
using OfficeMCP.Models;
using OfficeMCP.Services;
using System.ComponentModel;
using System.Text.Json;

namespace OfficeMCP.Tools;

/// <summary>
/// MCP Tools for creating and manipulating Excel workbooks.
/// </summary>
[McpServerToolType]
public sealed class ExcelDocumentTools(IExcelDocumentService excelService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    [McpServerTool, Description("Create a new Excel workbook (.xlsx) with an optional initial sheet name.")]
    public string CreateExcelWorkbook(
        [Description("Full file path for the new workbook (e.g., C:/Documents/data.xlsx)")] string filePath,
        [Description("Name for the first sheet (default: 'Sheet1')")] string? sheetName = null)
    {
        var result = excelService.CreateWorkbook(filePath, sheetName);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a new sheet to an existing Excel workbook.")]
    public string AddSheetToExcel(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name for the new sheet")] string sheetName)
    {
        var result = excelService.AddSheet(filePath, sheetName);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Set a value in a specific cell with optional formatting.")]
    public string SetExcelCellValue(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Cell reference (e.g., 'A1', 'B5', 'AA100')")] string cellReference,
        [Description("Value to set (numbers, text, or dates)")] string value,
        [Description("Make text bold")] bool bold = false,
        [Description("Make text italic")] bool italic = false,
        [Description("Enable text wrapping")] bool wrapText = false)
    {
        var formatting = new ExcelCellFormatting(Bold: bold, Italic: italic, WrapText: wrapText);
        var result = excelService.SetCellValue(filePath, sheetName, cellReference, value, formatting);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Set values in a range of cells starting from a specific cell.")]
    public string SetExcelRangeValues(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Starting cell reference (e.g., 'A1')")] string startCell,
        [Description("Values as JSON 2D array, e.g., [[\"A1\",\"B1\"],[\"A2\",\"B2\"]]")] string valuesJson)
    {
        try
        {
            var values = JsonSerializer.Deserialize<string[][]>(valuesJson, JsonOptions);
            if (values == null || values.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty values array"), JsonOptions);
            }

            var result = excelService.SetRangeValues(filePath, sheetName, startCell, values);
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description("Add a formatted table to an Excel sheet with optional auto-filter and styling.")]
    public string AddTableToExcel(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Starting cell reference for the table (e.g., 'A1')")] string startCell,
        [Description("Table data as JSON 2D array where first row can be headers")] string tableDataJson,
        [Description("First row contains headers")] bool hasHeaders = true)
    {
        try
        {
            var tableData = JsonSerializer.Deserialize<string[][]>(tableDataJson, JsonOptions);
            if (tableData == null || tableData.Length == 0)
            {
                return JsonSerializer.Serialize(new DocumentResult(false, "Invalid or empty table data"), JsonOptions);
            }

            var result = excelService.AddTable(filePath, sheetName, startCell, tableData, hasHeaders);
            return JsonSerializer.Serialize(result, JsonOptions);
        }
        catch (JsonException ex)
        {
            return JsonSerializer.Serialize(new DocumentResult(false, $"Invalid JSON format: {ex.Message}"), JsonOptions);
        }
    }

    [McpServerTool, Description("Add an image to an Excel sheet at a specific cell location.")]
    public string AddImageToExcel(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Full path to the image file")] string imagePath,
        [Description("Cell reference where image should be anchored (e.g., 'A1')")] string cellReference,
        [Description("Image width in inches")] double widthInches = 4.0,
        [Description("Image height in inches")] double heightInches = 3.0,
        [Description("Alt text for accessibility")] string? altText = null)
    {
        var options = new ImageOptions(
            WidthEmu: (long)(widthInches * 914400),
            HeightEmu: (long)(heightInches * 914400),
            AltText: altText
        );
        
        var result = excelService.AddImage(filePath, sheetName, imagePath, cellReference, options);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Merge a range of cells into a single cell.")]
    public string MergeExcelCells(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Starting cell of the merge range (e.g., 'A1')")] string startCell,
        [Description("Ending cell of the merge range (e.g., 'C1')")] string endCell)
    {
        var result = excelService.MergeCells(filePath, sheetName, startCell, endCell);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Set the width of a specific column.")]
    public string SetExcelColumnWidth(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Column index (1 = A, 2 = B, etc.)")] int columnIndex,
        [Description("Column width in character units")] double width)
    {
        var result = excelService.SetColumnWidth(filePath, sheetName, columnIndex, width);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Set the height of a specific row.")]
    public string SetExcelRowHeight(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Row index (1-based)")] int rowIndex,
        [Description("Row height in points")] double height)
    {
        var result = excelService.SetRowHeight(filePath, sheetName, rowIndex, height);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Add a formula to a cell. The formula will be calculated when opened in Excel.")]
    public string AddExcelFormula(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Cell reference for the formula (e.g., 'C1')")] string cellReference,
        [Description("Excel formula without the leading '=' (e.g., 'SUM(A1:B1)', 'AVERAGE(A1:A10)')")] string formula)
    {
        var result = excelService.AddFormula(filePath, sheetName, cellReference, formula);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Get the value of a specific cell.")]
    public string GetExcelCellValue(
        [Description("Path to the Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Cell reference (e.g., 'A1')")] string cellReference)
    {
        var result = excelService.GetCellValue(filePath, sheetName, cellReference);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Get values from a range of cells as tab-separated text.")]
    public string GetExcelRangeValues(
        [Description("Path to the Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName,
        [Description("Starting cell of the range (e.g., 'A1')")] string startCell,
        [Description("Ending cell of the range (e.g., 'C10')")] string endCell)
    {
        var result = excelService.GetRangeValues(filePath, sheetName, startCell, endCell);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Get all text content from a specific sheet.")]
    public string GetExcelSheetText(
        [Description("Path to the Excel workbook")] string filePath,
        [Description("Name of the sheet")] string sheetName)
    {
        var result = excelService.GetSheetText(filePath, sheetName);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Get all text content from all sheets in the workbook.")]
    public string GetExcelAllSheetsText(
        [Description("Path to the Excel workbook")] string filePath)
    {
        var result = excelService.GetAllSheetsText(filePath);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Delete a sheet from the workbook.")]
    public string DeleteExcelSheet(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Name of the sheet to delete")] string sheetName)
    {
        var result = excelService.DeleteSheet(filePath, sheetName);
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    [McpServerTool, Description("Rename a sheet in the workbook.")]
    public string RenameExcelSheet(
        [Description("Path to the existing Excel workbook")] string filePath,
        [Description("Current name of the sheet")] string oldName,
        [Description("New name for the sheet")] string newName)
    {
        var result = excelService.RenameSheet(filePath, oldName, newName);
        return JsonSerializer.Serialize(result, JsonOptions);
    }
}
