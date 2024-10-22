using ETS.DataImporter.DbContexts;
using ETS.DataImporter.DbContexts.Entities;
using Microsoft.EntityFrameworkCore;

namespace ETS.DataImporter.Seed;
public static class SeedPetkim
{
    public static async Task SeedDataAsync(ApplicationDbContext context)
    {

        // First, seed the Excel path data
        await SeedPathAsync(context);

        // Then, seed the formula data
        await SeedFormulaAsync(context);
    }

   private static async Task SeedPathAsync(ApplicationDbContext context)
{
    // Check if any path already exists
    if (!context.ExcelPathMasters.Any())
    {
        Console.WriteLine("Seeding Petkim Path");

        var pathMaster = new List<ExcelPathMaster>
        {
            new ExcelPathMaster
            {
                GMExcel_Name = "PETKIM Gross",
                GMExcel_Path = "/Users/suleyman/Desktop/Gross/", // Your specific path
                GMExcel_Description = "Petkim Gross Profit Forecast File",
                GMExcel_Freq = 1,  // 1=Hourly, 2=Daily, 3=Weekly, 4=Monthly
            }
        };

        // Now, we are adding the path master to the context
        await context.ExcelPathMasters.AddRangeAsync(pathMaster);
        await context.SaveChangesAsync();

        Console.WriteLine("Petkim path seeded successfully.");
    }
}

    private static async Task SeedFormulaAsync(ApplicationDbContext context)
    {
        var pathMaster = await context.ExcelPathMasters
            .FirstOrDefaultAsync(p => p.GMExcel_Name == "PETKIM Gross");


        // Check if pathMaster is null
        if (pathMaster == null)
        {
            Console.WriteLine("Error: 'PETKIM Forecast' PathMaster not found. Cannot seed formulas.");
            return;  // Early exit to avoid further null reference issues
        }


        // Check if the data already exists to avoid duplicate seeding
        if (!context.ExcelFormulaMasters.Any())
        {
            Console.WriteLine("Seeding Petkim Formulas");

            var petkimData = new List<ExcelFormulaMaster>
           {
            new ExcelFormulaMaster { GM_EXCEL_PRODUCT_NAME = "Product_Name", GM_EXCEL_COLUMN = "A1", GM_EXCEL_FORMULA = "Product_Name", GM_EXCEL_PATH_MASTER_ID = pathMaster.GMExcel_ID },
            new ExcelFormulaMaster { GM_EXCEL_PRODUCT_NAME = "Product_Forecasted_Value", GM_EXCEL_COLUMN = "B1", GM_EXCEL_FORMULA = "Product_Forecasted_Value", GM_EXCEL_PATH_MASTER_ID = pathMaster.GMExcel_ID },
            
        };



            Console.WriteLine("Seeding the database with Petkim data...");

            // Add the seed data to the database context
            await context.ExcelFormulaMasters.AddRangeAsync(petkimData);

            // Save the changes to the database
            await context.SaveChangesAsync();
            Console.WriteLine("Petkim Data seeded successfully.");

        }
        else
        {
            Console.WriteLine("Petkim data is already present.");
        }
    }
}
