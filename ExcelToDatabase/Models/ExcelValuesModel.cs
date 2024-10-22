namespace ETS.DataImporter.Models
{
    public class ExcelValuesModel
    {
        public int PathMasterId { get; set; }
        public string? ProductName { get; set; }
        public float Value { get; set; }
        public DateTime InsertedDate { get; set; }
        public DateTime DateTime { get; set; }
    }
}
