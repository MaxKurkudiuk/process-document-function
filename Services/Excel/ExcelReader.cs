using ProcessDocumentFunction.Models.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ProcessDocumentFunction.Services.Excel;

public static class ExcelReader
{
  public static IEnumerable<ReportLogOutData> ReadSheets(WorkbookPart workbookPart)
  {
    try
    {
      var allSheets = workbookPart.Workbook.Sheets?.OfType<Sheet>() ?? [];
      return allSheets
        .Where(s => s.Name?.Value != null)
        .Select(s => new ReportLogOutData()
        {
          SheetName = s.Name!.Value!,
          IsStats = s.Name!.Value!.EndsWith(" stats", StringComparison.CurrentCultureIgnoreCase)
        });
    }
    catch (Exception e)
    {
      throw new Exception("Input file is not valid", e);
    }
  }

  public static void ReadRawSheetsData(WorkbookPart workbookPart, IEnumerable<ReportLogOutData> sheetsInfo)
  {
    var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>() ?? [];
    var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;

    var tempList = sheetsInfo.ToList();

    foreach (var sheetInfo in tempList)
    {
      var sheet = sheets.FirstOrDefault(s => s.Name?.Value == sheetInfo.SheetName);
      if (sheet == null) continue;

      var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
      sheetInfo.RawData = ReadAllRows(worksheetPart, sharedStringTable);

      ExcelService.DetectHeaders(sheetInfo);
    }
  }

  private static object[,] ReadAllRows(WorksheetPart worksheetPart, SharedStringTable? sharedStringTable)
  {
    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
    var rows = sheetData?.Elements<Row>().ToList() ?? [];

    if (rows.Count == 0) return new object[0, 0];

    var maxRow = rows.Max(r => (int)r.RowIndex!.Value);
    var maxCol = rows.Max(r => r.Elements<Cell>()
        .Select(c => ExcelService.GetColumnIndex(c.CellReference?.Value ?? ""))
        .DefaultIfEmpty(0).Max()) + 1;

    var data = new object[maxRow, maxCol];

    foreach (var row in rows)
    {
      int rowIdx = (int)row.RowIndex!.Value - 1;

      foreach (var cell in row.Elements<Cell>())
      {
        var colIdx = ExcelService.GetColumnIndex(cell.CellReference?.Value ?? "");

        if (colIdx >= 0 && colIdx < maxCol)
        {
          data[rowIdx, colIdx] = ExcelService.GetCellValue(cell, sharedStringTable);
        }
      }
    }

    return data;
  }
}
