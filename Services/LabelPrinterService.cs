using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;
using InosentAnlageAufbauTool.Helpers;
using InosentAnlageAufbauTool.Models;

namespace InosentAnlageAufbauTool.Services
{
    /// <summary>
    /// Prints labels by writing the selected items into a small Excel file that GoLabel reads.
    /// </summary>
    public class LabelPrinterService
    {
        private readonly ExcelService _excelService;
        private readonly ILogger _logger;

        private const string SheetName = "Tabelle1";
        private readonly string _sensorExcelPath = Path.Combine(Path.GetTempPath(), "PowerAutomateGodexSensorDe.xlsx");
        private readonly string _ledExcelPath = Path.Combine(Path.GetTempPath(), "PowerAutomateGodexLightDe.xlsx");

        // Live settings (editable via Config)
        public string SensorPrinterIp { get; private set; } = "";
        public string LedPrinterIp { get; private set; } = "";
        public string SensorTemplatePath { get; private set; } = "";
        public string LedTemplatePath { get; private set; } = "";
        public string GoLabelExePath { get; private set; } = @"C:\Program Files (x86)\GoDEX\GoLabel II\GoLabel.exe";
        public int CopiesEach { get; private set; } = 2; // fixed at 2

        public LabelPrinterService(ExcelService excelService, ConfigService configService, ILogger logger)
        {
            _excelService = excelService;
            _logger = logger;

            var cfg = configService.Load();
            ApplySettings(cfg);
        }

        public void ApplySettings(ConfigSettings cfg)
        {
            SensorPrinterIp = cfg.SensorPrinterIp ?? "";
            LedPrinterIp = cfg.LedPrinterIp ?? "";
            SensorTemplatePath = cfg.SensorTemplatePath ?? "";
            LedTemplatePath = cfg.LedTemplatePath ?? "";
            GoLabelExePath = string.IsNullOrWhiteSpace(cfg.GoLabelExePath)
                ? GoLabelExePath
                : cfg.GoLabelExePath;
            CopiesEach = cfg.CopiesEach < 1 ? 2 : cfg.CopiesEach;
        }

        // Legacy: print from full Excel copies (kept for compatibility)
        public bool PrintSensorLabels(string srcPath)
        {
            var okCopy = _excelService.CopySensorData(srcPath, _sensorExcelPath);
            if (!okCopy)
            {
                _logger.Log("[Print] SENSOR: Excel konnte nicht erstellt/kopiert werden.");
                return false;
            }

            return RunGoLabel(SensorTemplatePath, SensorPrinterIp, "SENSOR", _sensorExcelPath);
        }

        public bool PrintLedLabels(string srcPath)
        {
            var okCopy = _excelService.CopyLedData(srcPath, _ledExcelPath);
            if (!okCopy)
            {
                _logger.Log("[Print] LED: Excel konnte nicht erstellt/kopiert werden.");
                return false;
            }

            return RunGoLabel(LedTemplatePath, LedPrinterIp, "LED", _ledExcelPath);
        }

        // New: print directly from selected in-memory models (preserves address/type/location/timeout)
        public bool PrintSensorAddresses(IEnumerable<Sensor> sensors)
            => PrintSensorsInternal(sensors, SensorTemplatePath, SensorPrinterIp, _sensorExcelPath, "SENSOR");

        public bool PrintLedAddresses(IEnumerable<Led> leds)
            => PrintLedsInternal(leds, LedTemplatePath, LedPrinterIp, _ledExcelPath, "LED");

        // Backward compatibility: accept plain address lists (fills minimal data)
        public bool PrintSensorAddresses(IEnumerable<int> addresses)
        {
            var list = new List<Sensor>();
            foreach (var addr in addresses) list.Add(new Sensor { Address = addr });
            return PrintSensorAddresses(list);
        }

        public bool PrintLedAddresses(IEnumerable<int> addresses)
        {
            var list = new List<Led>();
            foreach (var addr in addresses) list.Add(new Led { Address = addr });
            return PrintLedAddresses(list);
        }

        private bool PrintSensorsInternal(IEnumerable<Sensor> sensors, string template, string ip, string excelPath, string tag)
        {
            var list = new List<Sensor>(sensors ?? Array.Empty<Sensor>());
            if (list.Count == 0)
            {
                _logger.Log($"[Print] {tag}: keine Adressen ausgewaehlt.");
                return false;
            }

            try
            {
                WriteSensorsExcel(excelPath, list);
                return RunGoLabel(template, ip, tag, excelPath);
            }
            catch (Exception ex)
            {
                _logger.Log($"[Print] {tag}: {ex.Message}");
                return false;
            }
        }

        private bool PrintLedsInternal(IEnumerable<Led> leds, string template, string ip, string excelPath, string tag)
        {
            var list = new List<Led>(leds ?? Array.Empty<Led>());
            if (list.Count == 0)
            {
                _logger.Log($"[Print] {tag}: keine Adressen ausgewaehlt.");
                return false;
            }

            try
            {
                WriteLedsExcel(excelPath, list);
                return RunGoLabel(template, ip, tag, excelPath);
            }
            catch (Exception ex)
            {
                _logger.Log($"[Print] {tag}: {ex.Message}");
                return false;
            }
        }

        private void WriteSensorsExcel(string path, List<Sensor> sensors)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? Path.GetTempPath());
            using var package = new ExcelPackage(new FileInfo(path));
            var ws = package.Workbook.Worksheets[SheetName] ?? package.Workbook.Worksheets.Add(SheetName);
            ws.Cells.Clear();

            // Header (aligns with GAS sheet structure)
            ws.Cells[1, 1].Value = "Model";
            ws.Cells[1, 2].Value = "Address";
            ws.Cells[1, 3].Value = "Location";
            ws.Cells[1, 4].Value = "Buzzer";

            int row = 2;
            foreach (var s in sensors)
            {
                ws.Cells[row, 1].Value = s.Model ?? string.Empty;
                ws.Cells[row, 2].Value = s.Address;
                ws.Cells[row, 3].Value = s.Location ?? string.Empty;
                ws.Cells[row, 4].Value = string.Equals(s.Buzzer, "Disable", StringComparison.OrdinalIgnoreCase)
                    ? "Buzzer Disable"
                    : "Buzzer Enable";
                row++;
            }

            package.Save();
        }

        private void WriteLedsExcel(string path, List<Led> leds)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? Path.GetTempPath());
            using var package = new ExcelPackage(new FileInfo(path));
            var ws = package.Workbook.Worksheets[SheetName] ?? package.Workbook.Worksheets.Add(SheetName);
            ws.Cells.Clear();

            ws.Cells[1, 1].Value = "Model";
            ws.Cells[1, 2].Value = "Address";
            ws.Cells[1, 3].Value = "Location";
            ws.Cells[1, 4].Value = "Timeout";

            int row = 2;
            foreach (var led in leds)
            {
                ws.Cells[row, 1].Value = led.Model ?? string.Empty;
                ws.Cells[row, 2].Value = led.Address;
                ws.Cells[row, 3].Value = led.Location ?? string.Empty;
                ws.Cells[row, 4].Value = led.TimeOut;
                row++;
            }

            package.Save();
        }

        private bool RunGoLabel(string template, string ip, string tag, string excelPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ip))
                    throw new Exception($"{tag}: Printer IP:Port ist leer.");

                if (!File.Exists(GoLabelExePath))
                    throw new FileNotFoundException($"{tag}: GoLabel.exe nicht gefunden.", GoLabelExePath);

                if (!File.Exists(template))
                    throw new FileNotFoundException($"{tag}: Template nicht gefunden.", template);

                if (!File.Exists(excelPath))
                    throw new FileNotFoundException($"{tag}: Excel nicht gefunden.", excelPath);

                var args = string.Join(" ", new[]
                {
                    "-f", Quote(template),
                    "-c", CopiesEach.ToString(),
                    "-i", Quote(ip)
                });

                var psi = new ProcessStartInfo
                {
                    FileName = GoLabelExePath,
                    Arguments = args,
                    UseShellExecute = true,
                    CreateNoWindow = false
                };

                using var p = Process.Start(psi) ?? throw new Exception($"{tag}: Prozessstart fehlgeschlagen.");
                p.WaitForExit(10_000);
                if (p.ExitCode != 0)
                    throw new Exception($"{tag}: GoLabel ExitCode {p.ExitCode}.");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Log($"[Print] {tag}: {ex.Message}");
                return false;
            }
        }

        private static string Quote(string s) => $"\"{s}\"";
    }
}
