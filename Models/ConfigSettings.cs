namespace InosentAnlageAufbauTool.Models
{
    public class ConfigSettings
    {
        public string SensorPrinterIp { get; set; } = "";
        public string LedPrinterIp { get; set; } = "";
        public string SensorTemplatePath { get; set; } = "";
        public string LedTemplatePath { get; set; } = "";

        // optional: path to GoLabel.exe if you ever make that configurable
        public string GoLabelExePath { get; set; } = @"C:\Program Files (x86)\GoDEX\GoLabel II\GoLabel.exe";

        // copies fixed at 2 (as requested); keep field if you want to expose later
        public int CopiesEach { get; set; } = 2;
    }
}
