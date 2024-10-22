namespace ETS.DataImporter.Models
{
    public class ExcelFormulaMasterModel
    {
        public int FormulaId { get; set; }
        public string? ProductName { get; set; }
        public string? Column { get; set; }
        public string? Formula { get; set; }
        public int PathMasterId { get; internal set; }
    }
}
