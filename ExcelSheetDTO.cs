namespace AddClientTestCodes.DTO
{
    public class ExcelSheetDTO
    {
        public string PanelName { get; set; } = "";
        public int GeneCount { get; set; }
        public List<string>? GenesInPanel { get; set; }
        public string IncompleteTestCode { get; set; } = "";
        public string CustomTestCode { get; set; } = "";
    }
}