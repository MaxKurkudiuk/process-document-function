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

    foreach (var sheetInfo in sheetsInfo)
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

  public static int SearchForHeaders(IEnumerable<string> headers, ReportLogOutData baseSheet, bool startsWith, bool contains, out List<string> missingHeaders)
  {
    missingHeaders = [];
    var data = baseSheet.RawData;
    if (data.Length == 0) return -1;

    int columnsCount = data.GetLength(1);
    int rowsCount = data.GetLength(0);
    int headersRowIndex = -1;

    for (int rowIndex = 0; rowIndex < rowsCount; rowIndex++)
    {
      List<string> missingHeadersRowScope = [];
      List<string> rowValues = [];

      for (int columnIndex = 0; columnIndex < columnsCount; columnIndex++)
      {
        rowValues.Add(Convert.ToString(data[rowIndex, columnIndex]) ?? "");
      }

      foreach (string header in headers)
      {
        if (!rowValues.Any(rowValue => ExcelService.IsRowValueMatchHeader(rowValue, header, startsWith, contains)))
        {
          missingHeadersRowScope.Add(header);
        }
      }

      if (missingHeadersRowScope.Count == 0)
      {
        headersRowIndex = rowIndex;
        missingHeaders.Clear();
        break;
      }
      else
      {
        if (missingHeaders.Count == 0 || missingHeadersRowScope.Count < missingHeaders.Count)
        {
          missingHeaders = missingHeadersRowScope;
        }
      }
    }

    return headersRowIndex;
  }
}
