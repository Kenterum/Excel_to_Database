using System.Globalization;
using ETS.DataImporter.DbContexts.Entities;
using ETS.DataImporter.Models;
using OfficeOpenXml;
using Microsoft.EntityFrameworkCore;
using ETS.DataImporter.DbContexts;

public class ExcelService : IExcelService
{
    private readonly ApplicationDbContext _dbContext;
    private readonly List<DateTime> _monitoredMonths;

    public ExcelService(ApplicationDbContext dbContext)
    {
        _dbContext = dbContext;
        _monitoredMonths = new List<DateTime>(); // Listenin başlangıçta boş olarak oluşturulması

    }
      
    public async Task<Dictionary<string, List<Dictionary<string, object>>>> ReadExcelAsync(string filePath)
    {
        var result = new Dictionary<string, List<Dictionary<string, object>>>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            await package.LoadAsync(new FileInfo(filePath));
            var selectedSheets = package.Workbook.Worksheets;

            foreach (var worksheet in selectedSheets)
            {
                var sheetData = new List<Dictionary<string, object>>();
                var columnNames = Enumerable.Range(1, worksheet.Dimension.End.Column)
                                            .Select(col => worksheet.Cells[1, col].Value?.ToString()?.Trim() ?? $"Column{col}")
                                            .ToList();

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var rowData = new Dictionary<string, object>();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var columnName = columnNames[col - 1];
                        var cellValue = worksheet.Cells[row, col].Value;
                        rowData.Add(columnName, cellValue);
                    }
                    sheetData.Add(rowData);
                }

                result.Add(worksheet.Name, sheetData);
            }
        }

        return result;
    }

    public async Task WriteExcelAsync(string filePath, List<ExcelFormulaMaster> formulas)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var fileInfo = new FileInfo(filePath);
        using (var package = new ExcelPackage(fileInfo))
        {
            var workbook = package.Workbook;
            var otDataWorksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "OT_Data");

            if (otDataWorksheet != null)
            {
                workbook.Worksheets.Delete(otDataWorksheet);
            }

            otDataWorksheet = workbook.Worksheets.Add("OT_Data");

            foreach (var formula in formulas)
            {
                if (formula == null || string.IsNullOrWhiteSpace(formula.GM_EXCEL_COLUMN) || formula.GM_EXCEL_FORMULA == null)
                {
                    continue;
                }

                var cell = otDataWorksheet.Cells[formula.GM_EXCEL_COLUMN];

                if (!formula.GM_EXCEL_FORMULA.StartsWith("="))
                {
                    cell.Value = formula.GM_EXCEL_FORMULA;
                    cell.Style.Font.Bold = true;
                }
                else
                {
                    cell.Formula = formula.GM_EXCEL_FORMULA;
                }

                cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            }

            var startCell = "A1";
            var endCell = $"B{otDataWorksheet.Dimension.End.Row}";
            var tableRange = otDataWorksheet.Cells[$"{startCell}:{endCell}"];
            tableRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            tableRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            tableRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            tableRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            otDataWorksheet.Column(1).AutoFit();
            otDataWorksheet.Column(2).AutoFit();

            otDataWorksheet.Calculate();

            await package.SaveAsync();
        }
    }
public DateTime ExtractDateTimeFromPath(string folderPath, int frequency)
{
    string[] pathParts = folderPath.Split(Path.DirectorySeparatorChar);
    string datePart = pathParts.Last();

    DateTime parsedDateTime;

    switch (frequency)
    {
        case 1: // Hourly
            parsedDateTime = DateTime.ParseExact(datePart, "yyyy-MM-dd HH-mm-ss", CultureInfo.InvariantCulture);
            break;
        case 2: // Daily
            parsedDateTime = DateTime.ParseExact(datePart, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            break;
        case 3: // Weekly
            string[] weekParts = pathParts[pathParts.Length - 2].Split('_');
            if (weekParts.Length < 2 || !weekParts[0].Equals("Week", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Invalid weekly folder structure.");
            }
            int weekNumber = int.Parse(weekParts[1]);
            DateTime firstDayOfMonth = new DateTime(
                DateTime.ParseExact(pathParts[pathParts.Length - 3], "MMMM yyyy", CultureInfo.InvariantCulture).Year,
                DateTime.ParseExact(pathParts[pathParts.Length - 3], "MMMM yyyy", CultureInfo.InvariantCulture).Month, 1);

            DateTime firstMonday = firstDayOfMonth;
            while (firstMonday.DayOfWeek != DayOfWeek.Monday)
            {
                firstMonday = firstMonday.AddDays(1);
            }

            DateTime weekMonday = firstMonday.AddDays((weekNumber - 1) * 7);
            parsedDateTime = weekMonday;
            break;
        case 4: // Monthly
            parsedDateTime = DateTime.ParseExact(datePart, "yyyy-MM", CultureInfo.InvariantCulture);
            break;
        default:
            throw new ArgumentException("Invalid frequency value");
    }

    // Ensure the DateTime is in UTC
    return DateTime.SpecifyKind(parsedDateTime, DateTimeKind.Utc);
}

public async Task CreateOrUpdateExcelFileAsync(ExcelPathMasterModel pathMaster)
{
    // Get the root directory path for the month
    string rootDirectoryPath = Path.Combine(pathMaster.Path, $"{pathMaster.Date:MMMM yyyy}"); // E.g., "August 2024"
    Directory.CreateDirectory(rootDirectoryPath); // Create the root directory for the month

    // Get the number of days in the month
    int daysInMonth = DateTime.DaysInMonth(pathMaster.Date.Year, pathMaster.Date.Month);

    // Loop through all the days in the month
    for (int day = 1; day <= daysInMonth; day++)
    {
        // Get the date for the current day in the loop
        DateTime currentDay = new DateTime(pathMaster.Date.Year, pathMaster.Date.Month, day);

        // Create the directory for the current day
        string dailyDirectoryPath = Path.Combine(rootDirectoryPath, $"{currentDay:yyyy-MM-dd}");
        Directory.CreateDirectory(dailyDirectoryPath);

        // Create hourly directories if the frequency is hourly
        if (pathMaster.Frequency == 1) // Hourly
        {
            for (int hour = 0; hour < 24; hour++)
            {
                string hourlyDirectoryPath = Path.Combine(dailyDirectoryPath, $"{currentDay:yyyy-MM-dd} {hour:D2}-00-00");
                Directory.CreateDirectory(hourlyDirectoryPath); // Create the hourly directory
            }
        }
        else if (pathMaster.Frequency == 2) // Daily
        {
            // Daily directory already created
        }
        else if (pathMaster.Frequency == 3) // Weekly
        {
            int weekNumber = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(currentDay, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            string weeklyDirectoryPath = Path.Combine(rootDirectoryPath, $"Week_{weekNumber}");
            Directory.CreateDirectory(weeklyDirectoryPath); // Create the weekly directory
        }
        else if (pathMaster.Frequency == 4) // Monthly
        {
            // Monthly directory already created as the root directory
        }
    }
}

    public string GetFileNameBasedOnFrequency(ExcelPathMasterModel pathMasterModel)
    {
        string fileName = pathMasterModel.Name;
        string datePart = "";

        switch (pathMasterModel.Frequency)
        {
            case 1: // Hourly
                datePart = pathMasterModel.Date.ToString("yyyy-MM-dd HH-mm-ss");
                break;
            case 2: // Daily
                datePart = pathMasterModel.Date.ToString("yyyy-MM-dd");
                break;
            case 3: // Weekly
                datePart = $"Week_{CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(pathMasterModel.Date, CalendarWeekRule.FirstDay, DayOfWeek.Monday)}";
                break;
            case 4: // Monthly
                datePart = pathMasterModel.Date.ToString("yyyy-MM");
                break;  
            default:
                throw new ArgumentException("Invalid frequency value");
        }

        return $"{fileName}_{datePart}.xlsx";
    }

  public void CheckDirectoryStructure(ExcelPathMasterModel pathMasterModel)
    {
        DateTime currentMonthStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

        // Eğer liste boşsa (program ilk çalıştığında), sadece mevcut ayı ekleriz
        if (_monitoredMonths.Count == 0)
        {
            _monitoredMonths.Add(currentMonthStart);
        }

        // Eğer yeni bir ay geldi ve listeye eklenmediyse, ekleriz
        if (!_monitoredMonths.Contains(currentMonthStart))
        {
            _monitoredMonths.Add(currentMonthStart);
        }

        // Tüm dinlenmesi gereken aylar için klasör yapısını kontrol ederiz
        foreach (var month in _monitoredMonths)
        {
            CheckMonthStructure(pathMasterModel, month);
        }
    }

private void CheckMonthStructure(ExcelPathMasterModel pathMasterModel, DateTime monthStartDate)
{
    // Root directory for the current month
    string rootDirectoryPath = Path.Combine(pathMasterModel.Path, $"{monthStartDate:MMMM yyyy}");
    if (!Directory.Exists(rootDirectoryPath))
    {
        Directory.CreateDirectory(rootDirectoryPath);
    }

    DateTime currentDate = monthStartDate;

    // Hourly Frequency
    if (pathMasterModel.Frequency == 1) // Hourly
    {
        while (currentDate.Month == monthStartDate.Month)
        {
            string dailyDirectoryPath = Path.Combine(rootDirectoryPath, $"{currentDate:yyyy-MM-dd}");
            if (!Directory.Exists(dailyDirectoryPath))
            {
                Directory.CreateDirectory(dailyDirectoryPath);
            }

            for (int hour = 0; hour < 24; hour++)
            {
                string hourlyDirectoryPath = Path.Combine(dailyDirectoryPath, $"{currentDate:yyyy-MM-dd} {hour:D2}-00-00");
                if (!Directory.Exists(hourlyDirectoryPath))
                {
                    Directory.CreateDirectory(hourlyDirectoryPath);
                }
            }

            currentDate = currentDate.AddDays(1);
        }
    }
    // Daily Frequency
    else if (pathMasterModel.Frequency == 2) // Daily
    {
        while (currentDate.Month == monthStartDate.Month)
        {
            string dailyDirectoryPath = Path.Combine(rootDirectoryPath, $"{currentDate:yyyy-MM-dd}");
            if (!Directory.Exists(dailyDirectoryPath))
            {
                Directory.CreateDirectory(dailyDirectoryPath);
            }
            currentDate = currentDate.AddDays(1);
        }
    }
    // Weekly Frequency
  else if (pathMasterModel.Frequency == 3) // Weekly
{
    int weekNumber = 1;
    while (currentDate.Month == monthStartDate.Month)
    {
        // Find the next Monday
        while (currentDate.DayOfWeek != DayOfWeek.Monday && currentDate.Month == monthStartDate.Month)
        {
            currentDate = currentDate.AddDays(1);
        }

        // If we've moved out of the month, stop creating weeks
        if (currentDate.Month != monthStartDate.Month)
            break;

        // Create the weekly directory based on the Monday we found
        string weeklyDirectoryPath = Path.Combine(rootDirectoryPath, $"{monthStartDate:MMMM-yyyy}-Week-{weekNumber}");
        if (!Directory.Exists(weeklyDirectoryPath))
        {
            Directory.CreateDirectory(weeklyDirectoryPath);
        }

        // Move to the next Monday
        currentDate = currentDate.AddDays(7);
        weekNumber++;
    }
}
    // Monthly Frequency
    else if (pathMasterModel.Frequency == 4) // Monthly
    {
        // Monthly directory is already created as the root directory
        // No further sub-directories are needed for the monthly frequency
    }
}

  public Task<List<ExcelValuesModel>> ReadOTDataValuesAsync(string filePath, int frequency, DateTime lastWriteTimeUtc)
{
    var results = new List<ExcelValuesModel>();

    // Extract datetime from the folder path
    string folderDateTime = Path.GetDirectoryName(filePath);
    DateTime parsedDateTime = ExtractDateTimeFromPath(folderDateTime, frequency);

    using (var package = new ExcelPackage(new FileInfo(filePath)))
    {
        var otDataWorksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "OT_Data");

        if (otDataWorksheet == null)
        {
            throw new Exception("OT_Data sheet not found.");
        }

        for (int row = 2; row <= otDataWorksheet.Dimension.End.Row; row++)
        {
            var productName = otDataWorksheet.Cells[row, 1].Text;
            var calculatedValueText = otDataWorksheet.Cells[row, 2].Text;

            if (!string.IsNullOrWhiteSpace(productName) && !string.IsNullOrWhiteSpace(calculatedValueText))
            {
                if (float.TryParse(calculatedValueText, out float calculatedValue))
                {
                    results.Add(new ExcelValuesModel
                    {
                        ProductName = productName,
                        Value = calculatedValue,
                        DateTime = parsedDateTime, // Date from the folder
                        InsertedDate = lastWriteTimeUtc // Last write time of the file
                    });
                }
                else
                {
                    Console.WriteLine($"Warning: Could not parse value '{calculatedValueText}' for product '{productName}' in row {row}.");
                }
            }
            else
            {
                Console.WriteLine($"Warning: Missing product name or value in row {row}.");
            }
        }
    }

    return Task.FromResult(results);
}

}