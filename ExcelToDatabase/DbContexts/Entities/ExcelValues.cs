namespace ETS.DataImporter.DbContexts.Entities;

public class ExcelValues
{
    public int? GM_EXCEL_VALUES_ID { get; set; }
    public string? GM_EXCEL_PRODUCT_NAME { get; set; }
    public float GM_EXCEL_VALUE { get; set; }
    public DateTime GM_EXCEL_DATETIME { get; set; }
    public DateTime GM_EXCEL_INSERTED_DATE { get; set; }

    // Relationship to ExcelPathMaster
    public int GM_EXCEL_PATH_MASTER_ID { get; set; }
    public ExcelPathMaster? ExcelPathMaster { get; set; }
}