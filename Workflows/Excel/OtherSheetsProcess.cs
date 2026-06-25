using ProcessDocumentFunction.Models.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Spreadsheet;
using ProcessDocumentFunction.Models.Constants.Excel;
using ProcessDocumentFunction.Services.Excel;
using ProcessDocumentFunction.Services;

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
      sheetInfo.HeaderColumnsDictionary = ExcelService.GetHeaderColumnIndexesRowScope(sheetInfo, headers);
      var vacationDataList = ExcelService.CollectVacationData(sheetInfo);
      int columnsCount = sheetInfo.RawData.GetLength(1);
      ProcessInvalidHours(vacationDataList, workbookPart, sheetInfo.SheetName, columnsCount);
      ProcessDuplicateRows(vacationDataList, workbookPart, sheetInfo.SheetName, columnsCount);
      ProcessIncorrectTitles(vacationDataList, workbookPart, sheetInfo.SheetName, columnsCount);

      _logger.LogInformation($"End process sheet: {sheetInfo.SheetName}");
    }
  }

  private void ProcessInvalidHours(List<VacationData> vacationDataList, WorkbookPart workbookPart, string sheetName, int columnsCount)
  {
    var groups = vacationDataList.GroupBy(x => (x.UserName, x.GetDate())).ToList();
    foreach (var group in groups)
    {
      if (group.Sum(x => x.GetTimeSpent()) > 10)
      {
        foreach (var vacationData in group)
        {
          _logger.LogWarning($"Sheet: {sheetName}. Incorrect Time Spent (more than 10): {group.Sum(x => x.GetTimeSpent())}. Row: {vacationData.RowIdx}");
          ExcelUpdater.SetRowColor(workbookPart, sheetName, (uint)vacationData.RowIdx, "F4B084", 1, (uint)columnsCount);
        }
      }
    }
  }

  private void ProcessDuplicateRows(List<VacationData> vacationDataList, WorkbookPart workbookPart, string sheetName, int columnsCount)
  {
    var groups = vacationDataList.GroupBy(x => (x.Title, x.GetDate(), x.UserName, x.TimeSpent)).ToList();
    foreach (var group in groups)
    {
      if (group.Count() > 1)
      {
        foreach (var vacationData in group)
        {
          _logger.LogWarning($"Sheet: {sheetName}. Duplicate. Row: {vacationData.RowIdx}");
          ExcelUpdater.SetRowColor(workbookPart, sheetName, (uint)vacationData.RowIdx, "FF0000", 1, (uint)columnsCount);
        }
      }
    }
  }

  private void ProcessIncorrectTitles(List<VacationData> vacationDataList, WorkbookPart workbookPart, string sheetName, int columnsCount)
  {
    foreach (var vacationData in vacationDataList.Where(x => OtherSheetsConfig.IncorrectWorkTitles.Any(word => x.Title.Contains(word))))
    {
      _logger.LogWarning($"Sheet: {sheetName}. Incorrect title (not work). Row: {vacationData.RowIdx}");
      ExcelUpdater.SetRowColor(workbookPart, sheetName, (uint)vacationData.RowIdx, "FF0000", 1, (uint)columnsCount);
    }
  }
}
