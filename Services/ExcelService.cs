using DocumentFormat.OpenXml.Spreadsheet;
using ProcessDocumentFunction.Models.Excel;
using ProcessDocumentFunction.Models.Constants.Excel;
using System.Text.RegularExpressions;

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

  public static string GetColumnLetter(int columnIndex)
  {
    int dividend = columnIndex + 1;
    string result = "";

    while (dividend > 0)
    {
      int modulo = (dividend - 1) % 26;
      result = (char)('A' + modulo) + result;
      dividend = (dividend - modulo) / 26;
    }

    return result;
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

  public static bool IsRowValueMatchHeader(string? rowValue, string header, bool startsWith, bool contains)
  {
    if (string.IsNullOrEmpty(rowValue)) return false;

    if (startsWith)
      return rowValue.StartsWith(header, StringComparison.OrdinalIgnoreCase);
    if (contains)
      return rowValue.Contains(header, StringComparison.OrdinalIgnoreCase);
    return rowValue.Equals(header, StringComparison.OrdinalIgnoreCase);
  }

  public static Dictionary<string, int> GetHeaderColumnIndexesRowScope(ReportLogOutData sheetInfo, IEnumerable<string> headers)
  {
    Dictionary<string, int> indexByHeaderDictionary = [];
    int columnsCount = sheetInfo.RawData.GetLength(1);
    for (int columnIndex = 0; columnIndex < columnsCount; columnIndex++)
    {
      foreach (string header in headers)
      {
        if (sheetInfo.RawData[sheetInfo.HeaderIndex, columnIndex] != null
            && Regex.Replace(header.Trim().ToLower(), @"\s", "") ==
              Regex.Replace(sheetInfo.RawData[sheetInfo.HeaderIndex, columnIndex].ToString()!.Trim().ToLower(), @"\s", ""))
        {
          if (!indexByHeaderDictionary.ContainsKey(header.Trim()))
          {
            indexByHeaderDictionary.Add(header.Trim(), columnIndex);
            break;
          }
        }
      }
    }
    return indexByHeaderDictionary;
  }

  public static List<VacationData> CollectVacationData(ReportLogOutData sheetInfo)
  {
    var vacationDataList = new List<VacationData>();
    var titleColumnIdx = sheetInfo.GetColumnNumberByName(OtherDataConfig.TitleColumnName);
    var dateColumnIdx = sheetInfo.GetColumnNumberByName(OtherDataConfig.DateColumnName);
    var userNameColumnIdx = sheetInfo.GetColumnNumberByName(OtherDataConfig.UserNameColumnName);
    var timeSpentColumnIdx = sheetInfo.GetColumnNumberByName(OtherDataConfig.TimeSpentColumnName);
    int rowsCount = sheetInfo.RawData.GetLength(0);
    int columnsCount = sheetInfo.RawData.GetLength(1);
    for (int row = 1; row < rowsCount; row++)
    {
      var titleValue = string.IsNullOrEmpty(sheetInfo.RawData[row, titleColumnIdx].ToString()) ? 
        "" : sheetInfo.RawData[row, titleColumnIdx].ToString()!.Trim();
      titleValue = new string([.. titleValue.ToLower().Where(x => char.IsLetter(x) || char.IsNumber(x))]);
      var dateValue = string.IsNullOrEmpty(sheetInfo.RawData[row, dateColumnIdx].ToString()) ? 
        "" : sheetInfo.RawData[row, dateColumnIdx].ToString()!.Trim();
      var userNameValue = string.IsNullOrEmpty(sheetInfo.RawData[row, userNameColumnIdx].ToString()) ? 
        "" : sheetInfo.RawData[row, userNameColumnIdx].ToString()!.Trim();
      userNameValue = new string([.. userNameValue.ToLower().Where(x => char.IsLetter(x) || char.IsNumber(x))]);
      string timeSpentValue = "";
      if (timeSpentColumnIdx > 0)
        timeSpentValue = string.IsNullOrEmpty(sheetInfo.RawData[row, timeSpentColumnIdx].ToString()) ? 
          "" : sheetInfo.RawData[row, timeSpentColumnIdx].ToString()!;
      vacationDataList.Add(new VacationData(row, titleValue, dateValue, userNameValue, timeSpentValue));
    }
    return vacationDataList;
  }
}
