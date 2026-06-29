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
      _logger.LogInformation("FormatingOnly. End process sheet: {SheetName}", sheetInfo.SheetName);
      return;
    }
    sheetInfo.HeaderColumnsDictionary = ExcelService.GetHeaderColumnIndexesRowScope(sheetInfo, headers);
    var vacationDataList = ExcelService.CollectVacationData(sheetInfo);
    int columnsCount = sheetInfo.RawData.GetLength(1);
    var groups = vacationDataList.GroupBy(x => (x.UserName, x.GetDate())).ToList();
    foreach (var group in groups)
    {
      if (group.Count() <= 1 && group.Count(x => !IsTitleCorrect(x.Title)) <= 0)
        continue;
      foreach (var vacationData in group)
      {
        if (group.Count() > 1)
        {
          _excelUpdater.SetRowColor(workbookPart, sheetInfo.SheetName, (uint)vacationData.RowIdx, "FF0000", 1, (uint)columnsCount);
          _logger.LogWarning("Sheet: {SheetName}. Duplicate. Row: {RowIdx}", sheetInfo.SheetName, vacationData.RowIdx);
        }
        else if (!IsTitleCorrect(vacationData.Title))
        {
          _excelUpdater.SetRowColor(workbookPart, sheetInfo.SheetName, (uint)vacationData.RowIdx, "F4B084", 1, (uint)columnsCount);
          _logger.LogWarning("Sheet: {SheetName}. Incorrect Title: {Title}. Row: {RowIdx}", sheetInfo.SheetName, vacationData.Title, vacationData.RowIdx);
        }
      }
    }
    _logger.LogInformation("End process sheet: {SheetName}", sheetInfo.SheetName);
  }

  private bool IsTitleCorrect(string titleValue)
  {
    if (string.IsNullOrEmpty(titleValue)) return false;
    var lettersStr = new string([.. titleValue.ToLower().Where(char.IsLetter)]);
    if (lettersStr.Equals("vacation") || lettersStr.Equals("illness") || lettersStr.Equals("dayoff") || lettersStr.Equals("sickday"))
      return true;
    return false;
  }
}
