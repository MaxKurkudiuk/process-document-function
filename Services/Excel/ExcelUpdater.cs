using ProcessDocumentFunction.Models.Excel;
using ProcessDocumentFunction.Models.Constants.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;

namespace ProcessDocumentFunction.Services.Excel;

public class ExcelUpdater(ILogger<ExcelUpdater> logger)
{
  private readonly ILogger<ExcelUpdater> _logger = logger;

  public void ChangeSheetName(WorkbookPart workbookPart, IEnumerable<ReportLogOutData> sheetsInfo, string oldName, string newName)
  {
    var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>() ?? [];
    var sheet = sheets.FirstOrDefault(s => s.Name?.Value == oldName);
    if (sheet == null) return;

    sheet.Name = newName;

    var entry = sheetsInfo.FirstOrDefault(s => s.SheetName == oldName);
    entry?.SheetName = newName;
    _logger.LogInformation("Sheet name changed from '{oldName}' to '{newName}'", oldName, newName);
  }

  public static void SetHeadersFilter(ReportLogOutData sheetInfo, WorkbookPart workbookPart)
  {
    if (sheetInfo == null || sheetInfo.HeaderIndex < 0 || sheetInfo.RawData.Length == 0) return;

    var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>() ?? [];
    var sheet = sheets.FirstOrDefault(s => s.Name?.Value == sheetInfo.SheetName);
    if (sheet == null) return;

    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    var worksheet = worksheetPart.Worksheet;

    int headerRow = sheetInfo.HeaderIndex + 1;
    int lastColumn = sheetInfo.RawData.GetLength(1);
    string lastCol = ExcelService.GetColumnLetter(lastColumn - 1);

    var autoFilter = worksheet.GetFirstChild<AutoFilter>();
    if (autoFilter == null)
    {
      autoFilter = new AutoFilter();
      worksheet.Append(autoFilter);
    }
    autoFilter.Reference = $"A{headerRow}:{lastCol}{headerRow}";
  }

  public static void SetColumnsWidth(ReportLogOutData sheetInfo, WorkbookPart workbookPart)
  {
    if (sheetInfo?.RawData == null) return;
    int columnsCount = sheetInfo.RawData.GetLength(1);
    if (columnsCount == 0) return;

    var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>() ?? [];
    var sheet = sheets.FirstOrDefault(s => s.Name?.Value == sheetInfo.SheetName);
    if (sheet == null) return;

    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    var worksheet = worksheetPart.Worksheet;

    var columns = worksheet.GetFirstChild<Columns>();
    if (columns == null)
    {
      columns = new Columns();
      worksheet.InsertAt(columns, 0);
    }

    int maxCol = Math.Min(columnsCount, SheetSizeConfig.GetColumnsWidthInReqOrder().Length);
    for (int i = 0; i < maxCol; i++)
    {
      uint colNumber = (uint)(i + 1);
      RemoveColumn(columns, colNumber);
      columns.Append(new Column
      {
        Min = colNumber,
        Max = colNumber,
        Width = SheetSizeConfig.GetColumnsWidthInReqOrder()[i],
        CustomWidth = true
      });
    }
  }

  private static void RemoveColumn(Columns columns, uint colNumber)
  {
    var existing = columns.Elements<Column>().FirstOrDefault(c => c.Min?.Value == colNumber);
    if (existing != null)
      columns.RemoveChild(existing);
  }

  public void MoveSheetToTheEnd(WorkbookPart workbookPart, ReportLogOutData? sheetInfo)
  {
    if (string.IsNullOrEmpty(sheetInfo?.SheetName)) return;

    var sheets = workbookPart.Workbook.Sheets;
    if (sheets == null) return;

    var sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name?.Value == sheetInfo.SheetName);
    if (sheet == null) return;

    sheet.Remove();
    sheets.Append(sheet);

    _logger.LogInformation("Sheet '{sheetName}' moved to the end.", sheetInfo.SheetName);
  }
}
