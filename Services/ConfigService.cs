using InosentAnlageAufbauTool.Models;
using System;
using System.IO;
using System.Text.Json;

namespace InosentAnlageAufbauTool.Services
{
    public class ConfigService
    {
        private readonly string _configDir;
        private readonly string _configPath;

        public ConfigService()
        {
            _configDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "InosentAnlageAufbauTool");
            _configPath = Path.Combine(_configDir, "config.json");
        }

        public ConfigSettings Load()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    var json = File.ReadAllText(_configPath);
                    var cfg = JsonSerializer.Deserialize<ConfigSettings>(json) ?? new ConfigSettings();
                    return EnsureDefaults(cfg);
                }
            }
            catch { /* ignore; fall back to defaults */ }

            // Defaults point to your Resources folder
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            return new ConfigSettings
            {
                SensorTemplatePath = Path.Combine(baseDir, "Resources", "AutomaticSensorDE.ezpx"),
                LedTemplatePath = Path.Combine(baseDir, "Resources", "AutomaticLightDE.ezpx"),
                // IPs left empty by default; you'll set them once in the config dialog
                SensorPrinterIp = "10.1.40.87:9100",
                LedPrinterIp = "10.1.40.87:9100"
            };
        }

        public void Save(ConfigSettings settings)
        {
            Directory.CreateDirectory(_configDir);
            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_configPath, json);
        }

        private static ConfigSettings EnsureDefaults(ConfigSettings s)
        {
            // keep it simple; ensure CopiesEach >=1
            if (s.CopiesEach < 1) s.CopiesEach = 2;
            return s;
        }
    }
}
