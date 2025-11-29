using System.IO;

namespace InosentAnlageAufbauTool.Models
{
    public class ProjectContext
    {
        public int? Number { get; set; }
        public string ExcelPath { get; set; } = string.Empty;
        public string WorkingCopyPath { get; set; } = string.Empty;

        public string FindExcelPath(int nr)
        {
            var firstTwo = nr.ToString().Substring(0, 2);
            var basePath = $@"T:\INOSENT_Projekte\20{firstTwo}\{nr}\{nr}_Anlageinfos\DS_{nr}";
            foreach (var ext in new[] { ".xlsm", ".xlsx", ".xls" })
            {
                var path = Path.Combine(basePath, $"Liste_{nr}{ext}");
                if (File.Exists(path)) return path;
            }
            return string.Empty;
        }
    }
}
