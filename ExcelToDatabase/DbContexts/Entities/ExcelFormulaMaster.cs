namespace ETS.DataImporter.DbContexts.Entities;
public class ExcelFormulaMaster
{
    public int GM_EXCEL_FORMULA_ID { get; set; }  // Unique ID for each entry
    public string? GM_EXCEL_PRODUCT_NAME { get; set; }  // Product name remains the same for both
    public string? GM_EXCEL_COLUMN { get; set; }  // Only filled for Product Name entry
    public string?GM_EXCEL_FORMULA { get; set; }  // Only filled for Product Name entry

    public int GM_EXCEL_PATH_MASTER_ID { get; set; }  // Foreign key relationship
    public ExcelPathMaster? ExcelPathMaster { get; set; }  // Navigation property
}
