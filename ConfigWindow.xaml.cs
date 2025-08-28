using InosentAnlageAufbauTool.Models;
using InosentAnlageAufbauTool.Services;
using Microsoft.Win32;
using System.IO;
using System.Windows;

namespace InosentAnlageAufbauTool.Views
{
    public partial class ConfigWindow : Window
    {
        private readonly ConfigService _config;
        private ConfigSettings _settings = new();

        public ConfigWindow(ConfigService config)
        {
            InitializeComponent();
            _config = config;

            // Always load merged settings (saved + defaults)
            _settings = _config.Load();

            // Pre-fill UI
            txtSensorIP.Text = _settings.SensorPrinterIp ?? "";
            txtLedIP.Text = _settings.LedPrinterIp ?? "";
            txtTemplateSensor.Text = _settings.SensorTemplatePath ?? "";
            txtTemplateLED.Text = _settings.LedTemplatePath ?? "";
        }

        private void OnBrowseTemplateSensor(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Title = "Choose Sensor Template (.ezpx)",
                Filter = "GoLabel Template (*.ezpx)|*.ezpx",
                CheckFileExists = true
            };
            if (ofd.ShowDialog() == true) txtTemplateSensor.Text = ofd.FileName;
        }

        private void OnBrowseTemplateLED(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Title = "Choose LED Template (.ezpx)",
                Filter = "GoLabel Template (*.ezpx)|*.ezpx",
                CheckFileExists = true
            };
            if (ofd.ShowDialog() == true) txtTemplateLED.Text = ofd.FileName;
        }

        private void OnOK(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtSensorIP.Text) ||
                string.IsNullOrWhiteSpace(txtLedIP.Text) ||
                string.IsNullOrWhiteSpace(txtTemplateSensor.Text) ||
                string.IsNullOrWhiteSpace(txtTemplateLED.Text))
            {
                MessageBox.Show("Please fill in all fields.", "Validation",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!File.Exists(txtTemplateSensor.Text) || !File.Exists(txtTemplateLED.Text))
            {
                MessageBox.Show("One or more template files not found.", "Validation",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Persist back
            _settings.SensorPrinterIp = txtSensorIP.Text.Trim();
            _settings.LedPrinterIp = txtLedIP.Text.Trim();
            _settings.SensorTemplatePath = txtTemplateSensor.Text.Trim();
            _settings.LedTemplatePath = txtTemplateLED.Text.Trim();

            _config.Save(_settings);

            DialogResult = true;
            Close();
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
