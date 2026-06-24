using ProcessDocumentFunction.Models.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;

namespace ProcessDocumentFunction.Services.Excel;

public class ExcelValidator(ILogger<ExcelValidator> logger)
{
  private readonly ILogger<ExcelValidator> _logger = logger;

  public bool Validate(WorkbookPart workbookPart, IEnumerable<ReportLogOutData> sheetsInfo, string fileName)
  {
    try
    {
      foreach (var item in sheetsInfo)
      {
        _logger.LogInformation("Sheet name: {SheetName}", item.SheetName);
      }

      ExcelReader.ReadRawSheetsData(workbookPart, sheetsInfo);

      _logger.LogInformation("File {fileName} validation finished successfully", fileName);
      return sheetsInfo.Any();
    }
    catch
    {
      return false;
    }
  }
}
