using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ProcessDocumentFunction.Models.Constants.Excel;
using ProcessDocumentFunction.Models.Excel;
using ProcessDocumentFunction.Services;
using ProcessDocumentFunction.Services.Excel;

namespace ProcessDocumentFunction.Workflows.Excel;

public class VacationIllnesProcess(ILogger<VacationIllnesProcess> logger, ExcelUpdater excelUpdater)
{
  private readonly ILogger<VacationIllnesProcess> _logger = logger;
  private readonly ExcelUpdater _excelUpdater = excelUpdater;

  public void Execute(
    WorkbookPart workbookPart, 
    ReportLogOutData? sheetInfo, 
    bool isFormattingOnly)
  {
    if (string.IsNullOrEmpty(sheetInfo?.SheetName)) return;

    _logger.LogInformation("Start process sheet: {SheetName}", sheetInfo.SheetName);

    var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>() ?? [];
    var sheet = sheets.FirstOrDefault(s => s.Name?.Value == sheetInfo.SheetName);
    if (sheet == null) return;

    var headers = new string[3] { OtherDataConfig.TitleColumnName, OtherDataConfig.DateColumnName, OtherDataConfig.UserNameColumnName };
    sheetInfo.HeaderIndex = ExcelReader.SearchForHeaders(headers, sheetInfo, false, false, out List<string> missingHeaders);
    if (missingHeaders.Count != 0)
      throw new Exception($"There are missing required headers on the sheet {sheetInfo.SheetName} : {string.Join(", ", missingHeaders)}");

    ExcelUpdater.SetHeadersFilter(sheetInfo, workbookPart);
    ExcelUpdater.SetColumnsWidth(sheetInfo, workbookPart);
    if (isFormattingOnly)
    {
      _logger.LogInformation($"FormatingOnly. End process sheet: {sheetInfo.SheetName}");
      return;
    }
    sheetInfo.HeaderColumnsDictionary = ExcelService.GetHeaderColumnIndexesRowScope(sheetInfo, headers);


    _logger.LogInformation($"End process sheet: {sheetInfo.SheetName}");
  }
}
