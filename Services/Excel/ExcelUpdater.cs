using ProcessDocumentFunction.Models.Excel;
using ProcessDocumentFunction.Models.Constants.Excel;
using DocumentFormat.OpenXml;
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
      var sheetData = worksheet.GetFirstChild<SheetData>();
      if (sheetData != null)
        worksheet.InsertAfter(autoFilter, sheetData);
      else
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
      var sheetData = worksheet.GetFirstChild<SheetData>();
      if (sheetData != null)
        worksheet.InsertBefore(columns, sheetData);
      else
        worksheet.Append(columns);
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

  private Dictionary<string, uint> FormatData = [];

  private Stylesheet GenerateStylesheet(Stylesheet stylesheet, string colorHex, uint numberFormatId = 0)
  {
    stylesheet.Fills ??= new Fills();
    if (!stylesheet.Fills.Any())
    {
      stylesheet.Fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.None }));
      stylesheet.Fills.Append(new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }));
    }

    uint fillIndex = GetOrCreateFillIndex(stylesheet, colorHex);

    stylesheet.CellFormats ??= new CellFormats();
    var cellFormat = new CellFormat()
    {
      FontId = 0,
      FillId = fillIndex,
      BorderId = 0,
      FormatId = 0,
      ApplyFill = true
    };
    if (numberFormatId > 0)
    {
      cellFormat.NumberFormatId = numberFormatId;
      cellFormat.ApplyNumberFormat = true;
    }
    stylesheet.CellFormats.Append(cellFormat);
    uint formatIndex = (uint)stylesheet.CellFormats.Count() - 1;
    FormatData.TryAdd($"{colorHex}|{numberFormatId}", formatIndex);
    stylesheet.Save();
    return stylesheet;
  }

  private uint GetOrCreateFillIndex(Stylesheet stylesheet, string colorHex)
  {
    string rgbValue = $"FF{colorHex}";
    uint i = 0;
    foreach (var fill in stylesheet.Fills!.Elements<Fill>())
    {
      var rgbHex = fill.PatternFill?.ForegroundColor?.Rgb?.Value;
      if (rgbHex == rgbValue)
        return i;
      i++;
    }
    stylesheet.Fills.Append(new Fill(new PatternFill(
        new ForegroundColor() { Rgb = new HexBinaryValue() { Value = rgbValue } }
    )
    { PatternType = PatternValues.Solid }));
    return (uint)stylesheet.Fills.Count() - 1;
  }

  public void SetRowColor(
    WorkbookPart workbookPart,
    string sheetName,
    uint rowIndex,
    string colorHex,
    uint startCol,
    uint endCol)
  {
    rowIndex++;
    var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>() ?? [];
    var sheet = sheets.FirstOrDefault(s => s.Name?.Value == sheetName);
    if (sheet == null) return;

    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    var worksheet = worksheetPart.Worksheet;
    var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

    var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
    if (row == null) return;

    var stylesPart = workbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
    stylesPart ??= workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet ??= new Stylesheet();

    for (uint col = startCol; col <= endCol; col++)
    {
      string cellRef = $"{ExcelService.GetColumnLetter((int)col - 1)}{rowIndex}";
      var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRef);
      if (cell == null)
      {
        cell = new Cell { CellReference = cellRef };
        row.Append(cell);
      }

      uint existingNumberFormatId = 0;
      if (cell.StyleIndex?.Value != null)
      {
        var existingFormat = stylesPart.Stylesheet?.CellFormats?.Elements<CellFormat>()
            .ElementAtOrDefault((int)cell.StyleIndex.Value);
        if (existingFormat?.NumberFormatId?.Value > 0 && existingFormat.ApplyNumberFormat?.Value == true)
          existingNumberFormatId = existingFormat.NumberFormatId.Value;
      }

      string formatKey = $"{colorHex}|{existingNumberFormatId}";
      if (!FormatData.ContainsKey(formatKey))
      {
        GenerateStylesheet(stylesPart.Stylesheet!, colorHex, existingNumberFormatId);
      }
      cell.StyleIndex = FormatData[formatKey];
    }
    row.CustomFormat = true;

    if (row.Parent == null) sheetData.Append(row);
    stylesPart.Stylesheet!.Save();
  }
}
