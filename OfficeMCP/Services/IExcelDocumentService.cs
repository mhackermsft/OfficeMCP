using OfficeMCP.Models;

namespace OfficeMCP.Services;

/// <summary>
/// Service interface for Excel document operations.
/// </summary>
public interface IExcelDocumentService
{
    DocumentResult CreateWorkbook(string filePath, string? sheetName = null);
    DocumentResult AddSheet(string filePath, string sheetName);
    DocumentResult SetCellValue(string filePath, string sheetName, string cellReference, string value, ExcelCellFormatting? formatting = null);
    DocumentResult SetRangeValues(string filePath, string sheetName, string startCell, string[][] values);
    DocumentResult AddTable(string filePath, string sheetName, string startCell, string[][] data, bool hasHeaders = true);
    DocumentResult AddImage(string filePath, string sheetName, string imagePath, string cellReference, ImageOptions? options = null);
    DocumentResult MergeCells(string filePath, string sheetName, string startCell, string endCell);
    DocumentResult SetColumnWidth(string filePath, string sheetName, int columnIndex, double width);
    DocumentResult SetRowHeight(string filePath, string sheetName, int rowIndex, double height);
    DocumentResult AddFormula(string filePath, string sheetName, string cellReference, string formula);
    DocumentResult AutoFitColumn(string filePath, string sheetName, int columnIndex);
    DocumentResult ResizeTableToIncludeRange(string filePath, string sheetName, string startCell, string endCell);
    DocumentResult FormatCellRange(string filePath, string sheetName, string startCell, string endCell, ExcelCellFormatting formatting);
    ContentResult GetCellValue(string filePath, string sheetName, string cellReference);
    ContentResult GetRangeValues(string filePath, string sheetName, string startCell, string endCell);
    ContentResult GetSheetText(string filePath, string sheetName);
    ContentResult GetAllSheetsText(string filePath);
    DocumentResult DeleteSheet(string filePath, string sheetName);
    DocumentResult RenameSheet(string filePath, string oldName, string newName);
    
    // Formatting info methods
    ExcelRangeFormattingResult GetCellFormatting(string filePath, string sheetName, string cellReference);
    ExcelRangeFormattingResult GetRangeFormatting(string filePath, string sheetName, string startCell, string endCell);
}
