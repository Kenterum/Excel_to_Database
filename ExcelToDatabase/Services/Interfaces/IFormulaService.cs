using ETS.DataImporter.DbContexts.Entities;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ETS.DataImporter.Services.Interfaces
{
    public interface IFormulaService
    {
        // Get formulas from the database based on PathMasterId
        Task<List<ExcelFormulaMaster>> GetFormulasAsync(int pathMasterId);
    }
}
