namespace ETS.DataImporter.DbContexts.Entities
{
    public class ExcelPathMaster
    {
        public int GMExcel_ID { get; set; }
        public string? GMExcel_Name { get; set; }
        public string? GMExcel_Path { get; set; }
        public string? GMExcel_Description { get; set; }
        
        // New fields to handle file frequency and date
        public int GMExcel_Freq { get; set; }  // 1=Hourly, 2=Daily, 3=Weekly, 4=Monthly
        public DateTime GMExcel_Date { get; set; }  // The date derived from the filename
        
        public ICollection<ExcelFormulaMaster>? ExcelFormulaMasters { get; set; }
        public ICollection<ExcelValues>? ExcelValues { get; set; }
    }
}
