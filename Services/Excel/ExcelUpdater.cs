using ProcessDocumentFunction.Models.Excel;
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
    if (entry != null)
      entry.SheetName = newName;
    _logger.LogInformation("Sheet name changed from '{oldName}' to '{newName}'", oldName, newName);
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
