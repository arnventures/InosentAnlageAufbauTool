using System;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using OfficeOpenXml; // EPPlus
using InosentAnlageAufbauTool.Helpers;
using InosentAnlageAufbauTool.Services;
using InosentAnlageAufbauTool.ViewModels;

namespace InosentAnlageAufbauTool
{
    public partial class App : Application
    {
        private ServiceProvider? _services;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // EPPlus license (adjust if you have a commercial license)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var sc = new ServiceCollection();

            // Logging
            sc.AddSingleton<ILogger, Logger>();

            // Services
            sc.AddSingleton<ExcelService>();
            sc.AddSingleton<ConfigService>();
            sc.AddSingleton<LabelPrinterService>(); // needs ExcelService, ConfigService, ILogger
            sc.AddSingleton<SerialService>();        // takes ILogger (optional)

            // ViewModel + Window
            sc.AddSingleton<MainViewModel>();
            sc.AddSingleton<MainWindow>();           // ctor(MainViewModel)

            _services = sc.BuildServiceProvider();

            var main = _services.GetRequiredService<MainWindow>();
            main.Show();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            if (_services is IDisposable d)
            {
                try { d.Dispose(); } catch { /* ignore */ }
            }
            base.OnExit(e);
        }
    }
}
