using System;
using System.Diagnostics;
using System.IO;
using InosentAnlageAufbauTool.Helpers;
using InosentAnlageAufbauTool.Models;

namespace InosentAnlageAufbauTool.Services
{
    public class LabelPrinterService
    {
        private readonly ExcelService _excelService;
        private readonly ILogger _logger;

        // Live settings (editable via Config)
        public string SensorPrinterIp { get; private set; } = "";
        public string LedPrinterIp { get; private set; } = "";
        public string SensorTemplatePath { get; private set; } = "";
        public string LedTemplatePath { get; private set; } = "";
        public string GoLabelExePath { get; private set; } = @"C:\Program Files (x86)\GoDEX\GoLabel II\GoLabel.exe";
        public int CopiesEach { get; private set; } = 2; // fixed at 2 as requested

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

        public bool PrintLedLabels(string srcPath)
        {
            var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
            var dstPath = Path.Combine(baseDir, "GodexLightDe.xlsx");
            if (!_excelService.CopyLedData(srcPath, dstPath))
            {
                _logger.Log("[Print] LED: Excel konnte nicht erstellt/kopiert werden.");
                return false;
            }

            return RunGoLabel(LedTemplatePath, LedPrinterIp, "LED");
        }

        public bool PrintSensorLabels(string srcPath)
        {
            var baseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources");
            var dstPath = Path.Combine(baseDir, "GodexSensorDe.xlsx");
            if (!_excelService.CopySensorData(srcPath, dstPath))
            {
                _logger.Log("[Print] SENSOR: Excel konnte nicht erstellt/kopiert werden.");
                return false;
            }

            return RunGoLabel(SensorTemplatePath, SensorPrinterIp, "SENSOR");
        }

        private bool RunGoLabel(string template, string ip, string tag)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ip))
                    throw new Exception($"{tag}: Printer IP:Port ist leer.");

                if (!File.Exists(GoLabelExePath))
                    throw new FileNotFoundException($"{tag}: GoLabel.exe nicht gefunden.", GoLabelExePath);

                if (!File.Exists(template))
                    throw new FileNotFoundException($"{tag}: Template nicht gefunden.", template);

                var args = $"/P \"{template}\" /N {CopiesEach} /O \"tcp://{ip}\"";

                var psi = new ProcessStartInfo
                {
                    FileName = GoLabelExePath,
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var p = Process.Start(psi) ?? throw new Exception($"{tag}: Prozessstart fehlgeschlagen.");
                p.WaitForExit(10_000);
                if (p.ExitCode != 0)
                    throw new Exception($"{tag}: GoLabel ExitCode {p.ExitCode}.");
                _logger.Log($"[Print] {tag}: OK → {ip}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Log($"[Print] {tag}: {ex.Message}");
                return false;
            }
        }

    }
}
