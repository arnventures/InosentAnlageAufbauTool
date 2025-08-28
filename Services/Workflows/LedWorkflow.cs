using System;
using System;
using System.Threading;
using System.Threading.Tasks;
using InosentAnlageAufbauTool.Helpers;
using InosentAnlageAufbauTool.Services;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    public sealed class LedWorkflow : ILedWorkflow
    {
        private readonly SerialService _serial;
        private readonly ExcelService _excel;
        private readonly ILogger? _logger;

        public LedWorkflow(SerialService serial, ExcelService excel, ILogger? logger = null)
        {
            _serial = serial;
            _excel = excel;
            _logger = logger;
        }

        public async Task ConfigureFromExcelAsync(string excelPath, Action<string> log, CancellationToken token)
        {
            var leds = _excel.LoadLedData(excelPath);
            if (leds.Count == 0)
            {
                log("Keine LEDs zu programmieren (LIGHTIMOK leer oder keine Adressen > 185).");
                return;
            }

            foreach (var (address, timeout) in leds)
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    _serial.FlushBuffers();

                    // FC16: write registers 4..7: [Notbetrieb=0, Addr, Baud=9600, Key=0x8F8F]
                    ushort[] block = { 0, (ushort)address, 9600, 0x8F8F };
                    _serial.WriteMultiple(1, 4, block);
                    log($"LED @1: FC16 [4..7]=[0,{address},9600,0x8F8F]");

                    // Timeout at reg 3 (0 or 180)
                    await _serial.WriteSingleAsync(3, (ushort)timeout, 1);
                    log($"LED Timeout gesetzt (Reg3={timeout}).");

                    // Confirm alive (be lenient; LEDs don't reboot)
                    var t0 = DateTime.UtcNow;
                    bool ok = false;
                    while (!token.IsCancellationRequested && (DateTime.UtcNow - t0).TotalMilliseconds < 3000)
                    {
                        if (_serial.CheckFast((byte)address, 200, token)) { ok = true; break; }
                        try { await Task.Delay(150, token); } catch { }
                    }
                    if (!ok) log($"LED @ {address}: keine stabile Antwort – vermutlich ok.");
                    else log($"LED @ {address}: bestätigt.");
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    log($"LED {address}: Fehler: {ex.Message}");
                }
            }
        }
    }
}
