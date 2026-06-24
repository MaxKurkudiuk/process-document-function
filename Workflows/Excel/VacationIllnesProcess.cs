using Microsoft.Extensions.Logging;

namespace ProcessDocumentFunction.Workflows.Excel;

public class VacationIllnesProcess(ILogger<VacationIllnesProcess> logger)
{
  private readonly ILogger<VacationIllnesProcess> _logger = logger;

  public void Execute(string? sheetName)
  {
    if (string.IsNullOrEmpty(sheetName)) return;

    _logger.LogInformation($"Start process sheet: {sheetName}");
  }
}
