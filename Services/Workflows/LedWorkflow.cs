using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using InosentAnlageAufbauTool.Helpers;
using InosentAnlageAufbauTool.Models;
using InosentAnlageAufbauTool.Services;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    public sealed class LedWorkflow : ILedWorkflow
    {
        private readonly SerialService _serial;
        private readonly ILogger? _logger;

        private const int PollIntervalMs = 70;
        private const int StableWindowMs = 180;
        private const int PresenceTimeoutMs = 150;
        private const int RtuGapMs = 110;
        private const int VerifyTimeoutMs = 1500;

        private sealed record LedOutcome(bool Success, string? Note);

        public LedWorkflow(SerialService serial, ILogger? logger = null)
        {
            _serial = serial;
            _logger = logger;
        }

        public async Task ConfigureAsync(
            IList<Led> lights,
            Action<string> log,
            IProgress<LedProgress> progress,
            CancellationToken token)
        {
            foreach (var led in lights)
            {
                if (token.IsCancellationRequested) break;

                if (!led.IsSelected)
                {
                    progress.Report(new LedProgress(led, "Skipped", "nicht ausgewaehlt"));
                    log($"LED {led.Address} uebersprungen (nicht ausgewaehlt).");
                    continue;
                }

                progress.Report(new LedProgress(led, "Active", null));
                log($"LED {led.Index}: warte auf Geraet @1 ...");

                var ready = await WaitForAddr1Async(token);
                if (!ready)
                {
                    if (token.IsCancellationRequested) break;
                    progress.Report(new LedProgress(led, "Fail", "kein Geraet @1 gefunden"));
                    log($"LED {led.Index}: kein Geraet @1 gefunden.");
                    continue;
                }

                var outcome = await ConfigureSingleAsync(led, log, token);
                progress.Report(new LedProgress(led, outcome.Success ? "OK" : "Fail", outcome.Note));
                log(outcome.Success
                    ? $"LED {led.Address}: OK."
                    : $"LED {led.Address}: Fehler: {outcome.Note ?? "Unbekannt"}");
            }
        }

        private async Task<bool> WaitForAddr1Async(CancellationToken token)
        {
            int stableMs = 0;
            while (true)
            {
                if (token.IsCancellationRequested) return false;

                _serial.FlushBuffers();
                bool alive = IsAlive(1, token);

                if (alive)
                {
                    stableMs += PollIntervalMs;
                    if (stableMs >= StableWindowMs) return true;
                }
                else
                {
                    stableMs = 0;
                }

                try { await Task.Delay(PollIntervalMs, token); }
                catch (OperationCanceledException) { return false; }
            }
        }

        private async Task<LedOutcome> ConfigureSingleAsync(Led led, Action<string> log, CancellationToken token)
        {
            string? note = null;
            bool success = false;
            var targetAddr = (byte)led.Address;

            try
            {
                _serial.FlushBuffers();

                if (IsAlive(targetAddr, token))
                {
                    note = $"Adresse {targetAddr} ist bereits belegt.";
                    return new LedOutcome(false, note);
                }

                await RtuGapAsync(token);
                _serial.LedWriteAddressBaudWithKey(1, targetAddr, baud: 9600, mode: 0, key: 0x8F8F);
                log($"LED @1 -> Adresse {led.Address}, Baud 9600 (FC16).");

                await RtuGapAsync(token);
                var alive = await WaitAliveStableAsync(targetAddr, VerifyTimeoutMs, token);
                if (!alive)
                {
                    note = "Neue Adresse antwortet nicht stabil.";
                    return new LedOutcome(false, note);
                }

                try
                {
                    await _serial.WriteSingleAsync(3, (ushort)led.TimeOut, targetAddr);
                    log($"LED Timeout gesetzt (Reg3 @ {led.Address} = {led.TimeOut}).");
                }
                catch (Exception ex)
                {
                    log($"LED Timeout Hinweis: {ex.Message}");
                }

                success = true;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                note = ex.Message;
                _logger?.Log($"[LED] {ex.Message}");
                log($"LED {led.Address}: Fehler: {ex.Message}");
                success = false;
            }

            return new LedOutcome(success, note);
        }

        private async Task<bool> WaitAliveStableAsync(byte address, int timeoutMs, CancellationToken token)
        {
            var sw = Stopwatch.StartNew();
            int stableMs = 0;

            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (IsAlive(address, token))
                {
                    stableMs += PollIntervalMs;
                    if (stableMs >= StableWindowMs) return true;
                }
                else
                {
                    stableMs = 0;
                }

                try { await Task.Delay(PollIntervalMs, token); }
                catch (OperationCanceledException) { return false; }
            }

            return false;
        }

        private bool IsAlive(byte address, CancellationToken token)
        {
            try { return _serial.CheckFast(address, PresenceTimeoutMs, token); }
            catch (OperationCanceledException) { throw; }
            catch { /* retry with type read */ }

            try
            {
                var dt = _serial.ReadRegisterFast(address, SerialService.REG_DEVICE_TYPE, PresenceTimeoutMs, token);
                return dt > 0;
            }
            catch (OperationCanceledException) { throw; }
            catch { return false; }
        }

        private static Task RtuGapAsync(CancellationToken token)
            => Task.Delay(RtuGapMs, token);
    }
}
