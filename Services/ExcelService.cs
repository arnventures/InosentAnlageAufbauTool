using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace InosentAnlageAufbauTool.Services
{
    public class ExcelService
    {
        // EPPlus Lizenz einmalig setzen
        static ExcelService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private static readonly HashSet<string> LedSheets = new(StringComparer.OrdinalIgnoreCase)
        {
            "LIGHTIMOK","LIGHT","LEDIMOK","LED","Led","Light","Light 24V"
        };

        public List<(string Model, string? Location, int Address, bool DisableBuzzer)> LoadSensorData(string path)
        {
            var data = new List<(string, string?, int, bool)>();
            using var package = new ExcelPackage(new FileInfo(path));
            var ws = package.Workbook.Worksheets["GAS"];
            if (ws?.Dimension == null) return data;

            for (int row = 2; row <= ws.Dimension.End.Row; row++)
            {
                var model = ws.Cells[row, 1].Value?.ToString();
                var location = ws.Cells[row, 3].Value?.ToString();
                int addr = 0;
                _ = int.TryParse(ws.Cells[row, 2].Value?.ToString(), out addr);
                var disableBuzzer = string.Equals(ws.Cells[row, 4].Value?.ToString(), "Buzzer Disable", StringComparison.OrdinalIgnoreCase);

                if (!string.IsNullOrWhiteSpace(model) && addr > 0)
                    data.Add((model!, location, addr, disableBuzzer));
            }
            return data;
        }

        // LED-Daten aus Tabelle laden (B = Adresse, D = Timeout); nur Adressen > 185
        public List<(int Address, int Timeout)> LoadLedData(string path)
        {
            var result = new List<(int, int)>();
            using var package = new ExcelPackage(new FileInfo(path));
            var ws = package.Workbook.Worksheets.FirstOrDefault(s => s?.Name != null && LedSheets.Contains(s.Name.Trim()));
            if (ws?.Dimension == null) return result;

            for (int row = 2; row <= ws.Dimension.End.Row; row++)
            {
                if (int.TryParse(ws.Cells[row, 2].Value?.ToString(), out var addr) && addr > 185)
                {
                    int timeout = 0;
                    int.TryParse(ws.Cells[row, 4].Value?.ToString(), out timeout);
                    timeout = (timeout == 180) ? 180 : 0;
                    result.Add((addr, timeout));
                }
            }
            return result;
        }

        public bool CopyLedData(string srcPath, string dstPath)
        {
            try
            {
                if (!File.Exists(srcPath)) return false;

                using var srcPackage = new ExcelPackage(new FileInfo(srcPath));
                var srcSheet = srcPackage.Workbook.Worksheets
                    .FirstOrDefault(s => s?.Name != null && LedSheets.Contains(s.Name.Trim()));
                if (srcSheet == null || srcSheet.Dimension == null) return false;

                var dir = Path.GetDirectoryName(dstPath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir!);
                if (File.Exists(dstPath)) TryDelete(dstPath);

                using var dstPackage = new ExcelPackage();
                var dstSheet = dstPackage.Workbook.Worksheets.Add("ImportGodex");

                // Header kopieren (Zeile 1)
                for (int col = srcSheet.Dimension.Start.Column; col <= srcSheet.Dimension.End.Column; col++)
                    dstSheet.Cells[1, col].Value = srcSheet.Cells[1, col].Value;

                // Datenzeilen filtern (nur mit Sensor-ID > 0 in Spalte 2)
                int dstRow = 2;
                for (int row = srcSheet.Dimension.Start.Row + 1; row <= srcSheet.Dimension.End.Row; row++)
                {
                    if (int.TryParse(srcSheet.Cells[row, 2].Value?.ToString(), out var sid) && sid > 0)
                    {
                        for (int col = srcSheet.Dimension.Start.Column; col <= srcSheet.Dimension.End.Column; col++)
                            dstSheet.Cells[dstRow, col].Value = srcSheet.Cells[row, col].Value;
                        dstRow++;
                    }
                }

                if (dstRow <= 2) return false;

                dstPackage.SaveAs(new FileInfo(dstPath));
                return File.Exists(dstPath);
            }
            catch
            {
                return false;
            }
        }

        public bool CopySensorData(string srcPath, string dstPath)
        {
            try
            {
                if (!File.Exists(srcPath)) return false;

                using var srcPackage = new ExcelPackage(new FileInfo(srcPath));
                var srcSheet = srcPackage.Workbook.Worksheets["GAS"];
                if (srcSheet == null || srcSheet.Dimension == null) return false;

                var dir = Path.GetDirectoryName(dstPath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir!);
                if (File.Exists(dstPath)) TryDelete(dstPath);

                using var dstPackage = new ExcelPackage();
                var dstSheet = dstPackage.Workbook.Worksheets.Add("ImportGodex");

                // Header aus Quelle in Zeile 1
                for (int c = srcSheet.Dimension.Start.Column; c <= srcSheet.Dimension.End.Column; c++)
                    dstSheet.Cells[1, c].Value = srcSheet.Cells[1, c].Value;

                // Daten ab Zeile 2
                for (int r = 2; r <= srcSheet.Dimension.End.Row; r++)
                {
                    for (int c = 1; c <= srcSheet.Dimension.End.Column; c++)
                        dstSheet.Cells[r, c].Value = srcSheet.Cells[r, c].Value;
                }

                dstPackage.SaveAs(new FileInfo(dstPath));
                return File.Exists(dstPath);
            }
            catch
            {
                return false;
            }
        }

        // FIX: direkt auf Zeile = Index schreiben (keine +1 Verschiebung)
        public void WriteSerialToImport(string path, int index, int serial)
        {
            using var package = new ExcelPackage(new FileInfo(path));
            var ws = package.Workbook.Worksheets["Import"];
            if (ws == null) return;

            int row = index;          // index 1 -> row 1
            const int serialColumn = 5;
            ws.Cells[row, serialColumn].Value = serial;

            package.Save();
        }

        // Schneller: mehrere Seriennummern in einem Rutsch schreiben
        public void WriteSerialsBatch(string excelPath, List<(int RowIndex, int Serial)> rows)
        {
            if (rows == null || rows.Count == 0) return;

            using var package = new ExcelPackage(new FileInfo(excelPath));
            var ws = package.Workbook.Worksheets["Import"];
            if (ws == null) return;

            const int serialColumn = 5;
            foreach (var (rowIndex, serial) in rows)
            {
                ws.Cells[rowIndex, serialColumn].Value = serial;  // RowIndex ist 1-basiert
            }

            package.Save();
        }

        private static void TryDelete(string file)
        {
            try { File.Delete(file); } catch { /* ignore */ }
        }
    }
}
