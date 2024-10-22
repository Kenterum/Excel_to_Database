using ETS.DataImporter.Services;
using ETS.DataImporter.DbContexts;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using System.IO;
using ETS.DataImporter.Services.Interfaces;
using Microsoft.EntityFrameworkCore;
using ETS.DataImporter.Seed;
using Microsoft.Extensions.Logging;

namespace ETS.DataImporter.Program
{
    class Program
    {
        static async Task Main(string[] args) // Main method should be async to call async methods
        {
            // Set up configuration to read appsettings.json
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            // Build the host
            var host = Host.CreateDefaultBuilder(args)
                .ConfigureLogging(logging =>
                {
                    logging.ClearProviders();
                    logging.AddConsole();
                    logging.AddFilter("Microsoft.EntityFrameworkCore.Database.Command", LogLevel.Warning);
                })
                .ConfigureServices((context, services) =>
                {
                    // Register services for DI
                    services.AddSingleton<IExcelService, ExcelService>();
                    services.AddSingleton<IFormulaService, FormulaService>();
                    services.AddSingleton<IDatabaseService, DatabaseService>();

                    // Register ApplicationDbContext with the connection string from configuration
                    services.AddDbContext<ApplicationDbContext>(options =>
                        options.UseNpgsql(configuration.GetConnectionString("DefaultConnection")));

                    // Register the worker service
                    services.AddHostedService<FileProcessingWorker>();
                })
                .Build();

            // Seed the database before starting the worker service
            using (var scope = host.Services.CreateScope())
            {
                var dbContext = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
                await SeedPetkim.SeedDataAsync(dbContext); // Call your seed method here
            }

            // Run the worker service
            await host.RunAsync(); // Host needs to be awaited since it's async
        }
    }
}
