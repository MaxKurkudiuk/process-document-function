using ProcessDocumentFunction.Models.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Spreadsheet;
using ProcessDocumentFunction.Models.Constants.Excel;
using ProcessDocumentFunction.Services.Excel;

namespace ProcessDocumentFunction.Workflows.Excel;

public class OtherSheetsProcess(ILogger<OtherSheetsProcess> logger)
{
  private readonly ILogger<OtherSheetsProcess> _logger = logger;

  public void Execute(
    WorkbookPart workbookPart,
    IEnumerable<ReportLogOutData?> sheetsInfoToProcess,
    bool isFormattingOnly)
  {
    foreach (var sheetInfo in sheetsInfoToProcess)
    {
      if (string.IsNullOrEmpty(sheetInfo?.SheetName)) continue;
      _logger.LogInformation("Start process sheet: {SheetName}", sheetInfo.SheetName);

      var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>() ?? [];
      var sheet = sheets.FirstOrDefault(s => s.Name?.Value == sheetInfo.SheetName);
      if (sheet == null) continue;

      var headers = new string[4] { OtherDataConfig.TitleColumnName,
        OtherDataConfig.DateColumnName,
        OtherDataConfig.UserNameColumnName,
        OtherDataConfig.TimeSpentColumnName };
      sheetInfo.HeaderIndex = ExcelReader.SearchForHeaders(headers, sheetInfo, false, false, out List<string> missingHeaders);
      if (missingHeaders.Count != 0)
        throw new Exception($"There are missing required headers on the sheet {sheetInfo.SheetName} : {string.Join(", ", missingHeaders)}");

      ExcelUpdater.SetHeadersFilter(sheetInfo, workbookPart);
      ExcelUpdater.SetColumnsWidth(sheetInfo, workbookPart);
      if (isFormattingOnly)
      {
        _logger.LogInformation($"FormatingOnly. End process sheet: {sheetInfo.SheetName}");
        continue;
      }

      _logger.LogInformation($"End process sheet: {sheetInfo.SheetName}");
    }
  }
}
