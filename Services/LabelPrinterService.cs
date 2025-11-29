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
    /// Prints labels by writing the target address into a small Excel file that GoLabel reads.
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

        public bool PrintSensorLabels(string srcPath)
        {
            // Keep compatibility: prepare Excel from source, then print one label per row/address.
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

        public bool PrintSensorAddresses(IEnumerable<int> addresses)
        {
            return PrintAddressesInternal(addresses, SensorTemplatePath, SensorPrinterIp, _sensorExcelPath, "SENSOR");
        }

        public bool PrintLedAddresses(IEnumerable<int> addresses)
        {
            return PrintAddressesInternal(addresses, LedTemplatePath, LedPrinterIp, _ledExcelPath, "LED");
        }

        private bool PrintAddressesInternal(IEnumerable<int> addresses, string template, string ip, string excelPath, string tag)
        {
            var any = false;
            var allOk = true;
            foreach (var addrInt in addresses)
            {
                any = true;
                byte addr = (byte)addrInt;
                try
                {
                    WriteAddressExcel(excelPath, addr);
                    RunGoLabel(template, ip, tag, excelPath);
                    _logger.Log($"[Print] {tag} Adresse {addr}: OK");
                }
                catch (Exception ex)
                {
                    allOk = false;
                    _logger.Log($"[Print] {tag} Adresse {addr}: {ex.Message}");
                }
            }
            if (!any) _logger.Log($"[Print] {tag}: keine Adressen ausgewählt.");
            return any && allOk;
        }

        private void WriteAddressExcel(string path, byte address)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path) ?? Path.GetTempPath());
            using var package = new ExcelPackage(new FileInfo(path));
            var ws = package.Workbook.Worksheets[SheetName] ?? package.Workbook.Worksheets.Add(SheetName);
            ws.Cells.Clear();
            ws.Cells[1, 1].Value = address;
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
