using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ProcessDocumentFunction.Services.Excel;
using ProcessDocumentFunction.Models.Constants.Excel;
using ProcessDocumentFunction.Services;
using ProcessDocumentFunction.Workflows.Excel;

namespace ProcessDocumentFunction.Workflows;

public class ExcelWorkflow(
  ILogger<ExcelWorkflow> logger,
  ExcelValidator excelValidator,
  ExcelUpdater excelUpdater,
  VacationIllnesProcess vacationIllnesProcess,
  OtherSheetsProcess otherSheetsProcess
  )
{
  private readonly ILogger<ExcelWorkflow> _logger = logger;
  private readonly ExcelValidator _excelValidator = excelValidator;
  private readonly ExcelUpdater _excelUpdater = excelUpdater;
  private readonly VacationIllnesProcess _vacationIllnesProcess = vacationIllnesProcess;
  private readonly OtherSheetsProcess _otherSheetsProcess = otherSheetsProcess;

  public async Task<IActionResult> Execute(IFormFile file, bool isFormattingOnly)
  {
    using Stream fileStream = file.OpenReadStream();
    using var ms = new MemoryStream();
    await fileStream.CopyToAsync(ms);

    try
    {
      ms.Position = 0;
      using var doc = SpreadsheetDocument.Open(ms, true);
      var workbookPart = doc.WorkbookPart ?? throw new InvalidOperationException("No workbook part found.");

      var sheetsInfo = ExcelReader.ReadSheets(workbookPart).ToList();

      var isValid = _excelValidator.Validate(workbookPart, sheetsInfo, file.FileName);
      if (!isValid)
        throw new Exception("Input file is not valid");

      _excelUpdater.ChangeSheetName(workbookPart, sheetsInfo, AdminSheetNamesConfig.AdministrativeActivitiesDStatsOldName, AdminSheetNamesConfig.AdministrativeActivitiesDStatsNewName);
      _excelUpdater.ChangeSheetName(workbookPart, sheetsInfo, AdminSheetNamesConfig.AdministrativeActivitiesDOldName, AdminSheetNamesConfig.AdministrativeActivitiesDNewName);
      _excelUpdater.ChangeSheetName(workbookPart, sheetsInfo, AdminSheetNamesConfig.AdministrativeActivitiesHStatsOldName, AdminSheetNamesConfig.AdministrativeActivitiesHStatsNewName);
      _excelUpdater.ChangeSheetName(workbookPart, sheetsInfo, AdminSheetNamesConfig.AdministrativeActivitiesHOldName, AdminSheetNamesConfig.AdministrativeActivitiesHNewName);

      var vacationIllnessSheetName = ExcelService.GetVacationIllnessSheet(sheetsInfo);
      var vacationIllnessStatsSheetName = ExcelService.GetVacationIllnessStatsSheet(sheetsInfo);
      var otherSheetNames = ExcelService.GetOtherSheetNames(sheetsInfo);
      if (vacationIllnessSheetName == null)
        throw new Exception($"Fail to find any sheets with name starts with: {OtherDataConfig.VacationIllnessSheetKey}");
      _excelUpdater.MoveSheetToTheEnd(workbookPart, vacationIllnessStatsSheetName);
      _excelUpdater.MoveSheetToTheEnd(workbookPart, vacationIllnessSheetName);
      _vacationIllnesProcess.Execute(vacationIllnessSheetName);
      _otherSheetsProcess.Execute(otherSheetNames);

      workbookPart.Workbook.Save();
      doc.Save();

      ms.Position = 0;
      var resultFileName = $"{Path.GetFileNameWithoutExtension(file.FileName)}_result.xlsx";

      return new FileContentResult(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      {
        FileDownloadName = resultFileName
      };
    }
    catch (Exception ex)
    {
      _logger.LogError("OpenXML Error: {message}", ex.Message);
      return new StatusCodeResult(500);
    }
  }
}
