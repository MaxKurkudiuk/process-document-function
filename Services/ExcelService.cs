using DocumentFormat.OpenXml.Spreadsheet;
using ProcessDocumentFunction.Models.Excel;
using ProcessDocumentFunction.Models.Constants.Excel;

namespace ProcessDocumentFunction.Services;

public static class ExcelService
{
  public static int GetColumnIndex(string cellReference)
  {
    if (string.IsNullOrEmpty(cellReference)) return -1;

    var columnLetters = cellReference.TakeWhile(char.IsLetter).ToArray();
    if (columnLetters.Length == 0) return -1;

    return columnLetters.Aggregate(0, (a, c) => a * 26 + (c - 'A' + 1)) - 1;
  }

  public static string GetCellValue(Cell cell, SharedStringTable? sharedStringTable)
  {
    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
    {
      var index = int.Parse(cell.CellValue!.Text);
      return sharedStringTable?.ElementAt(index)?.InnerText ?? cell.CellValue.Text;
    }

    if (cell.DataType != null && cell.DataType.Value == CellValues.InlineString)
      return cell.InlineString?.Text?.Text ?? "";

    return cell.CellValue?.Text ?? "";
  }

  public static void DetectHeaders(ReportLogOutData sheetInfo)
  {
    if (sheetInfo.RawData.Length == 0) return;

    sheetInfo.HeaderIndex = 0;
    var columns = sheetInfo.RawData.GetLength(1);
    sheetInfo.Headers = [];
    sheetInfo.HeaderColumnsDictionary = [];

    for (int c = 0; c < columns; c++)
    {
      var header = sheetInfo.RawData[0, c]?.ToString() ?? "";
      sheetInfo.Headers.Add(header);

      if (!string.IsNullOrEmpty(header) && !sheetInfo.HeaderColumnsDictionary.ContainsKey(header))
        sheetInfo.HeaderColumnsDictionary[header] = c;
    }
  }

  public static ReportLogOutData? GetVacationIllnessSheet(IEnumerable<ReportLogOutData> sheetsInfo)
    => sheetsInfo
        .Where(x => !x.IsStats)
        .FirstOrDefault(x => x.SheetName.StartsWith(OtherDataConfig.VacationIllnessSheetKey, StringComparison.CurrentCultureIgnoreCase));

  public static ReportLogOutData? GetVacationIllnessStatsSheet(IEnumerable<ReportLogOutData> sheetsInfo)
    => sheetsInfo
        .Where(x => x.IsStats)
        .FirstOrDefault(x => x.SheetName.StartsWith(OtherDataConfig.VacationIllnessSheetKey, StringComparison.CurrentCultureIgnoreCase));
  
  public static IEnumerable<ReportLogOutData?> GetOtherSheetNames(IEnumerable<ReportLogOutData> sheetsInfo)
    => sheetsInfo
        .Where(x => !x.IsStats)
        .Where(x => !x.SheetName.StartsWith(OtherDataConfig.VacationIllnessSheetKey, StringComparison.CurrentCultureIgnoreCase));
}
