using Azure.Monitor.OpenTelemetry.Exporter;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Azure.Functions.Worker.OpenTelemetry;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ProcessDocumentFunction.Workflows;
using ProcessDocumentFunction.Services.Excel;
using ProcessDocumentFunction.Workflows.Excel;

var builder = FunctionsApplication.CreateBuilder(args);

builder.ConfigureFunctionsWebApplication();

if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("APPLICATIONINSIGHTS_CONNECTION_STRING")))
{
    builder.Services.AddOpenTelemetry()
        .UseFunctionsWorkerDefaults()
        .UseAzureMonitorExporter();
}

builder.Services.AddTransient<ExcelWorkflow>();
builder.Services.AddTransient<ExcelValidator>();
builder.Services.AddTransient<ExcelUpdater>();
builder.Services.AddTransient<VacationIllnesProcess>();
builder.Services.AddTransient<OtherSheetsProcess>();

builder.Build().Run();
