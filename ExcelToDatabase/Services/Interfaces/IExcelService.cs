using ETS.DataImporter.DbContexts.Entities;
using ETS.DataImporter.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

public interface IExcelService
{
    Task<Dictionary<string, List<Dictionary<string, object>>>> ReadExcelAsync(string filePath); // Reading data from Excel (M- OC and PETKIM)
    Task WriteExcelAsync(string filePath, List<ExcelFormulaMaster> formulas); // Writing formulas to OT_Data
    Task CreateOrUpdateExcelFileAsync(ExcelPathMasterModel pathMaster); // New method to handle file creation based on frequency
    
    // Ensure the signature matches the implementation in your class
    string GetFileNameBasedOnFrequency(ExcelPathMasterModel pathMasterModel); // Generate file name based on frequency and date

    // New method to ensure directory structure exists
    void CheckDirectoryStructure(ExcelPathMasterModel pathMasterModel); // Recreate missing directories

   public Task<List<ExcelValuesModel>> ReadOTDataValuesAsync(string filePath, int frequency, DateTime lastWriteTimeUtc);
        DateTime ExtractDateTimeFromPath(string folderPath, int frequency);

}
