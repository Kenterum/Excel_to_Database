using ETS.DataImporter.DbContexts;
using ETS.DataImporter.DbContexts.Entities;
using ETS.DataImporter.Services.Interfaces;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Threading.Tasks;

public class FormulaService : IFormulaService
{
    private readonly ApplicationDbContext _dbContext;

    // Constructor to inject the ApplicationDbContext
    public FormulaService(ApplicationDbContext dbContext)
    {
        _dbContext = dbContext;
    }

    // Retrieve formulas from the database based on the pathMasterId
    public async Task<List<ExcelFormulaMaster>> GetFormulasAsync(int pathMasterId)
    {
        return await _dbContext.ExcelFormulaMasters
                               .Where(f => f.GM_EXCEL_PATH_MASTER_ID == pathMasterId)
                               .ToListAsync();
    }
}
