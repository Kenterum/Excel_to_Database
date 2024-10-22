namespace ETS.DataImporter.Models
{
    public class ExcelPathMasterModel
    {
        public int PathMasterId { get; set; }
        public string? Name { get; set; }
        public string? Path { get; set; }
        public string? Description { get; set; }
        public int Frequency { get; set; }  // Frequency (1=Hourly, 2=Daily, 3=Weekly, 4=Monthly)
        public DateTime Date { get; set; }

    }
}
