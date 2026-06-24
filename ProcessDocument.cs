using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using ProcessDocumentFunction.Workflows;

namespace ProcessDocumentFunction;

public class ProcessDocument(ILogger<ProcessDocument> logger, ExcelWorkflow excelWorkflow)
{
  private readonly ILogger<ProcessDocument> _logger = logger;
  private readonly ExcelWorkflow _excelWorkflow = excelWorkflow;

  [Function("ProcessDocument")]
  public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
  {
    _logger.LogInformation("Processing incoming file request...");

    if (!req.HasFormContentType || req.Form.Files.Count == 0)
    {
      return new BadRequestObjectResult("Please upload a valid file.");
    }

    IFormFile file = req.Form.Files[0];

    if (!file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
    {
      return new BadRequestObjectResult("Only .xlsx files are supported.");
    }

    bool isFormattingOnly = false;
    if (req.Form.TryGetValue("isFormattingOnly", out var val))
      _ = bool.TryParse(val, out isFormattingOnly);

    return await _excelWorkflow.Execute(file, isFormattingOnly);
  }
}
