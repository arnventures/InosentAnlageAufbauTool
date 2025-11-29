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
        [ObservableProperty] private ObservableCollection<Led> lights = new();
        [ObservableProperty] private string projectNumber = string.Empty;
        [ObservableProperty] private string selectedPort = string.Empty;
        [ObservableProperty] private ObservableCollection<string> availablePorts = new();
        [ObservableProperty] private string statusText = "Bereit";
        [ObservableProperty] private string statusColor = "Yellow";
        [ObservableProperty] private string currentProjectLabel = "Current Anlage: -";
        [ObservableProperty] private string logText = string.Empty;
        [ObservableProperty] private bool isRunning;
        [ObservableProperty] private bool isComConnected;
        public bool AreSensorsComplete
        {
            get
            {
                var selected = Sensors.Where(s => s.IsSelected).ToList();
                return selected.Count > 0 && selected.All(s => s.Status is "OK" or "Skipped" or "Fail");
            }
        }
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
            _sensorWorkflow = new SensorWorkflow(_serialService, _logger);
            _ledWorkflow = new LedWorkflow(_serialService, _logger);
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

            IsComConnected = _serialService.IsConnected;
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
            IsComConnected = connected;
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
                _projectContext.WorkingCopyPath = _excelService.CreateWorkingCopy(_projectContext.ExcelPath);
                LoadFromWorkingCopy();
                CurrentProjectLabel = $"Current Anlage: {nr}";
                StatusText = "Excel geladen";
                StatusColor = "LightGreen";
                Log($"Excel geladen - {Sensors.Count} Sensoren, {Lights.Count} Leuchten (Arbeitskopie: {_projectContext.WorkingCopyPath}).");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excel konnte nicht geladen werden: {ex.Message}");
                Log($"[ERROR] Excel-Load: {ex.Message}");
            }
        }
        private void LoadFromWorkingCopy()
        {
            if (string.IsNullOrWhiteSpace(_projectContext.WorkingCopyPath))
                throw new InvalidOperationException("Keine Arbeitskopie gesetzt.");
            var sensorData = _excelService.LoadSensorData(_projectContext.WorkingCopyPath);
            Sensors.Clear();
            foreach (var data in sensorData)
            {
                Sensors.Add(new Sensor
                {
                    Index = data.RowIndex,
                    ExcelRow = data.RowIndex,
                    Model = data.Model,
                    Location = data.Location ?? string.Empty,
                    Address = data.Address,
                    Buzzer = data.DisableBuzzer ? "Disable" : "Enable",
                    Serial = null,
                    Status = "Pending",
                    IsSelected = true
                });
            }
            var ledData = _excelService.LoadLedData(_projectContext.WorkingCopyPath);
            Lights.Clear();
            foreach (var data in ledData)
            {
                Lights.Add(new Led
                {
                    Index = data.RowIndex,
                    Model = data.Model ?? string.Empty,
                    Location = data.Location ?? string.Empty,
                    Address = data.Address,
                    TimeOut = data.Timeout,
                    Status = "Pending",
                    IsSelected = true
                });
            }
            OnPropertyChanged(nameof(AreSensorsComplete));
        }
        // ---------- Run / Skip / Stop ----------
        // ---------- Run / Skip / Stop ----------
        // ---------- Run / Skip / Stop ----------
        [RelayCommand]
        private async Task Start()
        {
            if (string.IsNullOrEmpty(_projectContext.ExcelPath))
            {
                MessageBox.Show("Bitte zuerst Anlage laden.");
                return;
            }
            if (string.IsNullOrWhiteSpace(_projectContext.WorkingCopyPath))
            {
                try
                {
                    _projectContext.WorkingCopyPath = _excelService.CreateWorkingCopy(_projectContext.ExcelPath);
                    LoadFromWorkingCopy();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Arbeitskopie konnte nicht erstellt werden: {ex.Message}");
                    Log($"Excel-Arbeitskopie-Fehler: {ex.Message}");
                    return;
                }
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
            var selectedLights = Lights.Where(l => l.IsSelected).ToList();
            foreach (var s in selectedSensors)
            {
                s.Status = "Pending";
                s.Serial = null;
            }
            foreach (var l in selectedLights)
            {
                l.Status = "Pending";
            }
            _cts = new CancellationTokenSource();
            IsRunning = true;
            StatusText = "Sensoren...";
            StatusColor = "Yellow";
            var sensorProgress = new Progress<SensorProgress>(p =>
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
            var ledProgress = new Progress<LedProgress>(p =>
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (p.Led != null)
                    {
                        p.Led.Status = p.Status;
                    }
                });
            });
            try
            {
                var sensorResults = await _sensorWorkflow.RunAsync(
                    selectedSensors,
                    skipRequested: () =>
                    {
                        if (_skipEvent.WaitOne(0)) { _skipEvent.Reset(); return true; }
                        return false;
                    },
                    progress: sensorProgress,
                    log: Log,
                    token: _cts.Token);
                OnPropertyChanged(nameof(AreSensorsComplete));
                if (!_cts.IsCancellationRequested && sensorResults.Count > 0)
                {
                    try
                    {
                        _excelService.WriteSerialsBatch(
                            _projectContext.ExcelPath,
                            sensorResults.Select(r => (r.ExcelRowIndex, r.Serial)).ToList());
                        Log($"Excel-Update: {sensorResults.Count} Seriennummer(n) geschrieben.");
                    }
                    catch (Exception ex)
                    {
                        Log($"Excel-Batch-Fehler: {ex.Message}");
                    }
                }
                if (_cts.IsCancellationRequested)
                    return;
                StatusText = "Leuchten...";
                StatusColor = "Yellow";
                if (selectedLights.Any())
                {
                    await _ledWorkflow.ConfigureAsync(
                        selectedLights,
                        Log,
                        ledProgress,
                        _cts.Token);
                }
                else
                {
                    Log("Keine Leuchten ausgewählt.");
                }
            }
            catch (OperationCanceledException)
            {
                // Stop pressed
            }
            finally
            {
                IsRunning = false;
                StatusText = _cts is { IsCancellationRequested: true } ? "Abgebrochen" : "Fertig";
                StatusColor = _cts is { IsCancellationRequested: true } ? "Red" : "LightGreen";
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
            var addresses = Lights.Where(l => l.IsSelected).Select(l => l.Address).ToList();
            if (addresses.Count == 0)
            {
                MessageBox.Show("Keine Leuchten ausgew\u00e4hlt.");
                return;
            }
            var ok = _labelPrinterService.PrintLedAddresses(addresses);
            Log(ok ? $"LED-Labels gedruckt f\u00fcr {addresses.Count} Adressen." : "Fehler beim Drucken von LED-Labels.");
        }
        [RelayCommand]
        private void PrintSensorLabels()
        {
            var addresses = Sensors.Where(s => s.IsSelected).Select(s => s.Address).ToList();
            if (addresses.Count == 0)
            {
                MessageBox.Show("Keine Sensoren ausgew\u00e4hlt.");
                return;
            }
            var ok = _labelPrinterService.PrintSensorAddresses(addresses);
            Log(ok ? $"Sensor-Labels gedruckt f\u00fcr {addresses.Count} Adressen." : "Fehler beim Drucken von Sensor-Labels.");
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
