using ETS.DataImporter.DbContexts.Entities;
using ETS.DataImporter.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ETS.DataImporter.Services.Interfaces
{
    public interface IDatabaseService
    {
        Task SaveExcelModelsAsync(List<ExcelValuesModel> processedData);
        Task<List<ExcelValues>> GetProcessedExcelDataAsync(int pathMasterId);

        // New method for mapping raw data to ExcelValues
        List<ExcelValues> MapRawDataToExcelValues(List<Dictionary<string, object>> rawData, int pathMasterId);
    }
}
