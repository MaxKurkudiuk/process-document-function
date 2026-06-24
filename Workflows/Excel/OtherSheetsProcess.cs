using ProcessDocumentFunction.Models.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;

namespace ProcessDocumentFunction.Workflows.Excel;

public class OtherSheetsProcess(ILogger<OtherSheetsProcess> logger)
{
  private readonly ILogger<OtherSheetsProcess> _logger = logger;

  public void Execute(
    WorkbookPart workbookPart, 
    IEnumerable<ReportLogOutData?> sheetsInfoToProcess, 
    bool isFormattingOnly)
  {
    
  }
}
