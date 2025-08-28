using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using InosentAnlageAufbauTool.Helpers;
using InosentAnlageAufbauTool.Models;
using InosentAnlageAufbauTool.Services;
using InosentAnlageAufbauTool.Services.Workflows;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO.Ports;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace InosentAnlageAufbauTool.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly ExcelService _excelService;
        private readonly SerialService _serialService;
        private readonly LabelPrinterService _labelPrinterService;
        private readonly ConfigService _configService;
        private readonly ILogger _logger;
        private readonly ISensorWorkflow _sensorWorkflow;
        private readonly ILedWorkflow _ledWorkflow;

        private readonly ProjectContext _projectContext = new();

        private CancellationTokenSource? _cts;
        private readonly AutoResetEvent _skipEvent = new(false);

        [ObservableProperty] private ObservableCollection<Sensor> sensors = new();
        [ObservableProperty] private string projectNumber = string.Empty;
        [ObservableProperty] private string selectedPort = string.Empty;
        [ObservableProperty] private ObservableCollection<string> availablePorts = new();
        [ObservableProperty] private string statusText = "Bereit";
        [ObservableProperty] private string statusColor = "Yellow";
        [ObservableProperty] private string currentProjectLabel = "Current Anlage: –";
        [ObservableProperty] private string logText = string.Empty;
        [ObservableProperty] private bool isRunning;

        public bool AreSensorsComplete => Sensors.Count > 0 && Sensors.All(s => s.Status is "OK" or "Skipped");

        public MainViewModel(
            ExcelService excelService,
            SerialService serialService,
            LabelPrinterService labelPrinterService,
            ILogger logger,
            ConfigService configService)
        {
            _excelService = excelService;
            _serialService = serialService;
            _labelPrinterService = labelPrinterService;
            _configService = configService;
            _logger = logger;

            _sensorWorkflow = new SensorWorkflow(_serialService, _excelService, _logger);
            _ledWorkflow = new LedWorkflow(_serialService, _excelService, _logger);

            RefreshPortsCommand.Execute(null);
        }

        // Auto-connect whenever SelectedPort changes in the UI
        partial void OnSelectedPortChanged(string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                // Try connect automatically (no need to press "Verbinden")
                ConnectToPort();
            }
        }

        // ---------- Logging ----------
        private void Log(string msg)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                LogText += $"{DateTime.Now:HH:mm:ss}: {msg}\n";
            });
            _logger.Log(msg);
        }

        // ---------- Ports ----------
        [RelayCommand]
        private void RefreshPorts()
        {
            AvailablePorts.Clear();
            foreach (var p in SerialPort.GetPortNames()) AvailablePorts.Add(p);
            if (AvailablePorts.Any() && string.IsNullOrWhiteSpace(SelectedPort))
                SelectedPort = AvailablePorts[0];

            if (AvailablePorts.Any())
            {
                StatusText = "COM-Ports aktualisiert";
                StatusColor = "LightGreen";
                Log($"Verfügbare Ports: {string.Join(", ", AvailablePorts)}");

                // Auto-connect to the first available port if not already connected
                if (!_serialService.IsConnected && !string.IsNullOrWhiteSpace(SelectedPort))
                {
                    ConnectToPort();
                }
            }
            else
            {
                StatusText = "Kein COM-Port verfügbar";
                StatusColor = "Red";
                Log($"Verfügbare Ports: Keine");
            }
        }

        [RelayCommand]
        private void ConnectToPort()
        {
            if (string.IsNullOrWhiteSpace(SelectedPort))
            {
                StatusText = "Kein Port ausgewählt";
                StatusColor = "Red";
                return;
            }
            Log($"Versuche Verbindung zu {SelectedPort}...");
            var connected = _serialService.Connect(SelectedPort);
            StatusText = connected ? $"{SelectedPort}: connected" : $"{SelectedPort}: Verbindung fehlgeschlagen";
            StatusColor = connected ? "LightGreen" : "Red";
            Log($"Verbindung zu {SelectedPort}: {(connected ? "OK" : $"fehlgeschlagen ({_serialService.LastError})")})");
        }

        // ---------- Excel ----------
        [RelayCommand]
        private void LoadExcel()
        {
            if (!int.TryParse(ProjectNumber, out var nr) || nr <= 0)
            {
                MessageBox.Show("Bitte gültige Anlage-Nr. eingeben.");
                return;
            }
            _projectContext.Number = nr;
            _projectContext.ExcelPath = _projectContext.FindExcelPath(nr);
            if (string.IsNullOrEmpty(_projectContext.ExcelPath))
            {
                var ofd = new Microsoft.Win32.OpenFileDialog { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm" };
                if (ofd.ShowDialog() == true) _projectContext.ExcelPath = ofd.FileName;
                else return;
            }

            try
            {
                var sensorData = _excelService.LoadSensorData(_projectContext.ExcelPath);
                Sensors.Clear();
                for (int i = 0; i < sensorData.Count; i++)
                {
                    var data = sensorData[i];
                    Sensors.Add(new Sensor
                    {
                        Index = i + 1,
                        Model = data.Model,
                        Location = data.Location ?? string.Empty,
                        Address = data.Address,
                        Buzzer = data.DisableBuzzer ? "Disable" : "Enable",
                        Status = "Pending"
                    });
                }
                CurrentProjectLabel = $"Current Anlage: {nr}";
                Log($"Excel geladen – {Sensors.Count} Sensoren.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excel konnte nicht geladen werden: {ex.Message}");
                Log($"[ERROR] Excel-Load: {ex.Message}");
            }
        }

        // ---------- Run / Skip / Stop ----------
        [RelayCommand]
        private async Task Start()
        {
            if (string.IsNullOrEmpty(_projectContext.ExcelPath))
            {
                MessageBox.Show("Bitte zuerst Anlage laden.");
                return;
            }

            // If not connected yet but ports exist, try auto-connect once more
            if (!_serialService.IsConnected && !string.IsNullOrWhiteSpace(SelectedPort))
                ConnectToPort();

            if (!_serialService.IsConnected)
            {
                MessageBox.Show("COM-Port nicht verbunden.");
                return;
            }

            var selectedSensors = Sensors.Where(s => s.IsSelected).ToList();
            if (!selectedSensors.Any())
            {
                MessageBox.Show("Keine Sensoren ausgewählt.");
                return;
            }

            _cts = new CancellationTokenSource();
            IsRunning = true;
            StatusText = "Läuft …";
            StatusColor = "Yellow";

            var progress = new Progress<SensorProgress>(p =>
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (p.Sensor != null)
                    {
                        p.Sensor.Status = p.Status;
                        if (p.Serial.HasValue)
                            p.Sensor.Serial = p.Serial.Value;
                    }
                    OnPropertyChanged(nameof(AreSensorsComplete));
                });
            });

            try
            {
                await _sensorWorkflow.RunAsync(
                    selectedSensors,
                    _projectContext.ExcelPath,
                    skipRequested: () =>
                    {
                        if (_skipEvent.WaitOne(0)) { _skipEvent.Reset(); return true; }
                        return false;
                    },
                    progress: progress,
                    log: Log,
                    token: _cts.Token);
            }
            catch (OperationCanceledException) { /* Stop pressed */ }
            finally
            {
                IsRunning = false;
                StatusText = _cts.IsCancellationRequested ? "Abgebrochen" : "Fertig";
                StatusColor = _cts.IsCancellationRequested ? "Red" : "LightGreen";
                _cts?.Dispose();
                _cts = null;
            }
        }

        [RelayCommand]
        private void Skip() => _skipEvent.Set();

        [RelayCommand]
        private void Stop()
        {
            try { _cts?.Cancel(); } catch { }
        }

        // ---------- Printing ----------
        [RelayCommand]
        private void PrintLedLabels()
        {
            if (string.IsNullOrEmpty(_projectContext.ExcelPath))
            {
                MessageBox.Show("Bitte zuerst Anlage laden.");
                return;
            }
            var ok = _labelPrinterService.PrintLedLabels(_projectContext.ExcelPath);
            Log(ok ? "LED-Labels gedruckt." : $"Fehler beim Drucken von LED-Labels. {_labelPrinterService}");
        }

        [RelayCommand]
        private void PrintSensorLabels()
        {
            if (string.IsNullOrEmpty(_projectContext.ExcelPath))
            {
                MessageBox.Show("Bitte zuerst Anlage laden.");
                return;
            }
            var ok = _labelPrinterService.PrintSensorLabels(_projectContext.ExcelPath);
            Log(ok ? "Sensor-Labels gedruckt." : $"Fehler beim Drucken von Sensor-Labels. {_labelPrinterService}");
        }

        // ---------- Config ----------
        [RelayCommand]
        private void OpenConfig()
        {
            try
            {
                var wnd = new InosentAnlageAufbauTool.Views.ConfigWindow(_configService)
                {
                    Owner = Application.Current.MainWindow
                };

                if (wnd.ShowDialog() == true)
                {
                    _labelPrinterService.ApplySettings(_configService.Load());
                    Log("Konfiguration gespeichert.");
                    StatusText = "Konfiguration aktualisiert";
                    StatusColor = "LightGreen";
                }
            }
            catch (Exception ex)
            {
                Log($"Config-Fehler: {ex.Message}");
                StatusText = "Config-Fehler";
                StatusColor = "Red";
            }
        }

        // ---------- Help ----------
        [RelayCommand]
        private void ShowHelp()
        {
            MessageBox.Show(
                "Inosent Anlage Aufbau Tool\n\n" +
                "Ablauf:\n1. COM-Port verbinden.\n2. Anlage-Nr. eingeben → OK.\n3. Sensoren auswählen.\n4. Start – wartet auf Gerät @1, konfiguriert, schreibt Serien-Nr.\n5. Skip/Stop.\n\nMenü Print druckt Etiketten."
            );
        }
    }
}
