using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InosentAnlageAufbauTool.Helpers;
using InosentAnlageAufbauTool.Models;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    public sealed class LedWorkflow : ILedWorkflow
    {
        private readonly SerialService _serial;
        private readonly ILogger? _logger;

        private const int VerifyTimeoutMs = 1600;
        private const int PollDelayMs = 80;
        private const int PresenceTimeoutMs = 160;

        public LedWorkflow(SerialService serial, ILogger? logger = null)
        {
            _serial = serial;
            _logger = logger;
        }

        public async Task ConfigureAsync(
            System.Collections.Generic.IList<Led> lights,
            Action<string> log,
            IProgress<LedProgress> progress,
            CancellationToken token)
        {
            foreach (var led in lights)
            {
                token.ThrowIfCancellationRequested();
                progress.Report(new LedProgress(led, "Active", null));

                try
                {
                    _serial.FlushBuffers();

                    // FC16: write registers 4..7: [Notbetrieb=0, Addr, Baud=9600, Key=0x8F8F]
                    ushort[] block = { 0, (ushort)led.Address, 9600, 0x8F8F };
                    _serial.WriteMultiple(1, 4, block);
                    log($"LED @1: Adresse -> {led.Address}, Baud 9600");

                    await _serial.WriteSingleAsync(3, (ushort)led.TimeOut, 1);
                    log($"LED Timeout gesetzt (Reg3={led.TimeOut}).");

                    bool ok = await ConfirmAsync((byte)led.Address, token);
                    progress.Report(new LedProgress(led, ok ? "OK" : "Fail", null));
                    log(ok ? $"LED @ {led.Address}: best�tigt." : $"LED @ {led.Address}: keine stabile Antwort.");
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    progress.Report(new LedProgress(led, "Fail", ex.Message));
                    log($"LED {led.Address}: Fehler: {ex.Message}");
                    _logger?.Log($"[LED] {ex.Message}");
                }
            }
        }

        private async Task<bool> ConfirmAsync(byte address, CancellationToken token)
        {
            var start = DateTime.UtcNow;
            while ((DateTime.UtcNow - start).TotalMilliseconds < VerifyTimeoutMs)
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    if (_serial.CheckFast(address, PresenceTimeoutMs, token))
                        return true;
                }
                catch (OperationCanceledException) { throw; }
                catch { /* retry */ }

                try { await Task.Delay(PollDelayMs, token); } catch (OperationCanceledException) { throw; }
            }
            return false;
        }
    }
}
