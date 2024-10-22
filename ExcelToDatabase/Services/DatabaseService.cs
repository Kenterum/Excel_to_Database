using ETS.DataImporter.DbContexts;
using ETS.DataImporter.Services.Interfaces;
using ETS.DataImporter.DbContexts.Entities;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using System;
using ETS.DataImporter.Models;

public class DatabaseService : IDatabaseService
{
    private readonly ApplicationDbContext _context;

    public DatabaseService(ApplicationDbContext context)
    {
        _context = context;
    }

    // Save processed Excel values to the database
   public async Task SaveExcelModelsAsync(List<ExcelValuesModel> processedData)
    {
        // Map models to entities
        var excelEntities = processedData.Select(model => new ExcelValues
        {
            GM_EXCEL_PATH_MASTER_ID = model.PathMasterId,
            GM_EXCEL_PRODUCT_NAME = model.ProductName,
            GM_EXCEL_VALUE = model.Value,
            GM_EXCEL_DATETIME = DateTime.SpecifyKind(model.DateTime, DateTimeKind.Unspecified),
        GM_EXCEL_INSERTED_DATE = DateTime.SpecifyKind(model.InsertedDate, DateTimeKind.Unspecified)
        }).ToList();

        // Save to the database
        await _context.ExcelValues.AddRangeAsync(excelEntities);
        await _context.SaveChangesAsync();
    }

    // Retrieve processed Excel data from the database for a given PathMasterId
    public async Task<List<ExcelValues>> GetProcessedExcelDataAsync(int pathMasterId)
    {
        return await _context.ExcelValues
            .Where(ev => ev.GM_EXCEL_PATH_MASTER_ID == pathMasterId)
            .ToListAsync();
    }

    // Fetch formulas from the database based on PathMasterId
    public async Task<List<ExcelFormulaMaster>> GetFormulasAsync(int pathMasterId)
    {
        return await _context.ExcelFormulaMasters
            .Where(f => f.GM_EXCEL_PATH_MASTER_ID == pathMasterId)
            .ToListAsync();
    }

    // Method to map the raw data from Excel into ExcelValues
    public List<ExcelValues> MapRawDataToExcelValues(List<Dictionary<string, object>> rawData, int pathMasterId)
    {
        var excelValuesList = new List<ExcelValues>();

        foreach (var row in rawData)
        {
            var excelValue = new ExcelValues
            {
                GM_EXCEL_PRODUCT_NAME = row.ContainsKey("Product") ? row["Product"]?.ToString() : null,
                GM_EXCEL_VALUE = row.ContainsKey("Value") && float.TryParse(row["Value"]?.ToString(), out float value) ? value : 0f,
                GM_EXCEL_DATETIME = row.ContainsKey("Date") && DateTime.TryParse(row["Date"]?.ToString(), out DateTime date) ? date.ToUniversalTime() : DateTime.UtcNow,
                GM_EXCEL_INSERTED_DATE = DateTime.UtcNow,
                GM_EXCEL_PATH_MASTER_ID = pathMasterId
            };

            excelValuesList.Add(excelValue);
        }

        return excelValuesList;
    }
}
