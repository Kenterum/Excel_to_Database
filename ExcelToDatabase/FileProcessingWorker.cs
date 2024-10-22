using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using ETS.DataImporter.DbContexts;
using ETS.DataImporter.Models;
using Microsoft.EntityFrameworkCore;
using ETS.DataImporter.Services.Interfaces;

namespace ETS.DataImporter
{
    public class FileProcessingWorker : BackgroundService
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly ILogger<FileProcessingWorker> _logger;
        private readonly TimeZoneInfo _serverTimeZone;

        public FileProcessingWorker(IServiceProvider serviceProvider, ILogger<FileProcessingWorker> logger)
        {
            _serviceProvider = serviceProvider;
            _logger = logger;

            // Assuming server timezone is GMT+03:00
            _serverTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Europe/Istanbul"); // Replace with correct ID if needed
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("FileProcessingWorker is starting.");

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    using (var scope = _serviceProvider.CreateScope())
                    {
                        var dbContext = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
                        var excelService = scope.ServiceProvider.GetRequiredService<IExcelService>();
                        var formulaService = scope.ServiceProvider.GetRequiredService<IFormulaService>();
                        var dbService = scope.ServiceProvider.GetRequiredService<IDatabaseService>();

                        // Fetch all PathMaster records
                        var pathMasters = await dbContext.ExcelPathMasters.ToListAsync();

                        foreach (var pathMaster in pathMasters)
                        {
                            var pathMasterModel = new ExcelPathMasterModel
                            {
                                PathMasterId = pathMaster.GMExcel_ID,
                                Name = pathMaster.GMExcel_Name,
                                Path = pathMaster.GMExcel_Path,
                                Description = pathMaster.GMExcel_Description,
                                Frequency = pathMaster.GMExcel_Freq,
                                Date = DateTime.UtcNow
                            };

                            // Ensure folder structure is correct from the current month onward
                            excelService.CheckDirectoryStructure(pathMasterModel);

                            // Get all valid .xlsx files in the directory (no time filtering)
                            _logger.LogInformation($"Searching files in directory: {pathMasterModel.Path}");
                            var allFiles = Directory
                                .GetFiles(pathMasterModel.Path, "*.xlsx", SearchOption.AllDirectories)
                                .Select(f => new FileInfo(f))
                                .Where(f => !IsHiddenOrSystemFile(f))
                                .ToList();

                            _logger.LogInformation($"Found {allFiles.Count} files to process.");
                            foreach (var file in allFiles)
                            {
                                try
                                {
                                    _logger.LogInformation($"Processing file {file.Name}...");

                                    // Extract datetime from folder path
                                    string folderPath = Path.GetDirectoryName(file.FullName);
                                    DateTime fileDateTime = DateTime.SpecifyKind(excelService.ExtractDateTimeFromPath(folderPath, pathMasterModel.Frequency), DateTimeKind.Utc);

                                    // Retrieve last inserted date from the database, ensuring it's in UTC
                                    var lastInsertedRecord = await dbContext.ExcelValues
                                        .Where(e => e.GM_EXCEL_PATH_MASTER_ID == pathMaster.GMExcel_ID)
                                        .OrderByDescending(e => e.GM_EXCEL_INSERTED_DATE)
                                        .FirstOrDefaultAsync();

                                    DateTime? lastInsertedDate = lastInsertedRecord?.GM_EXCEL_INSERTED_DATE;
                                    DateTime? lastInsertedDateLocal = lastInsertedDate.HasValue
                                     ? TimeZoneInfo.ConvertTimeFromUtc(lastInsertedDate.Value, _serverTimeZone)
                                     : (DateTime?)null;

                                    // Convert file's last write time from UTC to server's local time
                                    DateTime fileLastWriteTimeUtc = file.LastWriteTimeUtc;
                                    DateTime fileLastWriteTimeLocal = TimeZoneInfo.ConvertTimeFromUtc(file.LastWriteTimeUtc, _serverTimeZone);
                                    fileLastWriteTimeUtc = DateTime.SpecifyKind(fileLastWriteTimeUtc, DateTimeKind.Utc);

                                    // Compare file's last write time with the last inserted date in the database
                                    if (lastInsertedDate == null || fileLastWriteTimeLocal > lastInsertedDateLocal)
                                    {
                                        var formulas = await formulaService.GetFormulasAsync(pathMaster.GMExcel_ID);
                                        await excelService.WriteExcelAsync(file.FullName, formulas);

                                        var calculatedValues = await excelService.ReadOTDataValuesAsync(file.FullName, pathMasterModel.Frequency, fileLastWriteTimeUtc);

                                        // Create a list of ExcelValuesModel
                                        var processedExcelValues = new List<ExcelValuesModel>();

                                        foreach (var value in calculatedValues)
                                        {
                                            var model = new ExcelValuesModel
                                            {
                                                PathMasterId = pathMaster.GMExcel_ID,
                                                ProductName = value.ProductName,
                                                Value = value.Value,
                                                DateTime = fileDateTime,
                                                InsertedDate = fileLastWriteTimeLocal // Save the file's last write time in local time
                                            };

                                            processedExcelValues.Add(model);
                                        }

                                        // Save the models (mapping to entity happens inside DatabaseService)
                                        await dbService.SaveExcelModelsAsync(processedExcelValues);
                                    }
                                    else
                                    {
                                        _logger.LogInformation($"File {file.Name} has not been modified since last processing. Skipping.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogError(ex, $"Error processing file {file.Name}: {ex.Message}");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "An error occurred while processing files.");
                }

                // Wait for the next interval (e.g., every 5 seconds for testing purposes)
                await Task.Delay(TimeSpan.FromSeconds(5), stoppingToken);
            }

            _logger.LogInformation("FileProcessingWorker is stopping.");
        }

        private bool IsHiddenOrSystemFile(FileInfo fileInfo)
        {
            return (fileInfo.Attributes & FileAttributes.Hidden) != 0 ||
                   (fileInfo.Attributes & FileAttributes.System) != 0;
        }
    }
}
