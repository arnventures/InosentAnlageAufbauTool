using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InosentAnlageAufbauTool.Helpers;
using InosentAnlageAufbauTool.Models;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    public sealed class SensorWorkflow : ISensorWorkflow
    {
        private readonly SerialService _serial;
        private readonly ExcelService _excel;
        private readonly ILogger? _logger;

        // Tuned timings (kept tight for speed)
        private const int PollFastMs = 120;
        private const int PollDelayMs = 120;
        private const int RebootWaitMs = 1200;
        private const int PostMoveTimeoutMs = 4000;

        public SensorWorkflow(SerialService serial, ExcelService excel, ILogger? logger = null)
        {
            _serial = serial;
            _excel = excel;
            _logger = logger;
        }

        /// <summary>
        /// Programs all sensors and writes serials to Excel once at the end (batch).
        /// Waits FOREVER for addr=1 (unless Stop or user Skip).
        /// </summary>
        public async Task RunAsync(
            IList<Sensor> sensors,
            string excelPath,
            Func<bool> skipRequested,
            IProgress<SensorProgress> progress,
            Action<string> log,
            CancellationToken token)
        {
            var batchSerials = new List<(int RowIndex, int Serial)>(); // 1-based row == sensor.Index
            Sensor? previous = null;

            foreach (var s in sensors)
            {
                token.ThrowIfCancellationRequested();

                progress.Report(new SensorProgress(s, "Active", null, null));
                previous = s;

                // Wait FOREVER for any device on addr=1 (only Skip/Stop break this)
                log("Warte auf Gerät @1 … (Skip überspringt; Stop beendet)");
                var hasDevice = await WaitForAddr1ForeverAsync(skipRequested, token);
                if (!hasDevice)
                {
                    // If we are here, token was canceled or user skipped explicitly
                    if (token.IsCancellationRequested) throw new OperationCanceledException(token);
                    progress.Report(new SensorProgress(s, "Skipped", null, "vom Benutzer übersprungen"));
                    log($"Sensor {s.Index} übersprungen.");
                    continue;
                }

                // Configure one sensor
                var (ok, serial) = await ConfigureSingleAsync(s, log, token);

                // Buffer serial for batch write (use fallback read @new address already inside ConfigureSingle)
                if (ok && serial.HasValue && serial.Value > 0)
                {
                    batchSerials.Add((s.Index, serial.Value)); // write to row == index (no +1)
                }

                progress.Report(new SensorProgress(s, ok ? "OK" : "Fail", serial, null));
                log(ok ? $"Sensor {s.Index} OK (SN={serial})" : $"Sensor {s.Index} Fail.");
            }

            if (previous != null)
                progress.Report(new SensorProgress(previous, previous.Status, previous.Serial, null));

            // Batch write serials only if run wasn't canceled
            if (!token.IsCancellationRequested && batchSerials.Count > 0)
            {
                try
                {
                    _excel.WriteSerialsBatch(excelPath, batchSerials);
                    log($"Excel-Update: {batchSerials.Count} Seriennummer(n) geschrieben.");
                }
                catch (Exception ex)
                {
                    _logger?.Log($"[SensorWorkflow] Excel-Batch-Fehler: {ex.Message}");
                    log($"Excel-Batch-Fehler: {ex.Message}");
                }
            }
        }

        // ---------- internals ----------

        /// <summary>
        /// Wait FOREVER for a device that answers at addr=1.
        /// Breaks only on Skip (returns false) or Stop (throws OperationCanceledException).
        /// </summary>
        private async Task<bool> WaitForAddr1ForeverAsync(
            Func<bool> skipRequested,
            CancellationToken token)
        {
            while (true)
            {
                token.ThrowIfCancellationRequested();
                if (skipRequested()) return false;

                try
                {
                    if (_serial.CheckFast(1, PollFastMs, token))
                        return true;
                }
                catch (OperationCanceledException) { throw; }
                catch { /* ignore transient errors */ }

                try { await Task.Delay(PollDelayMs, token); } catch (OperationCanceledException) { throw; }
            }
        }

        private async Task<(bool Success, int? Serial)> ConfigureSingleAsync(
            Sensor sensor,
            Action<string> log,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            try
            {
                _serial.FlushBuffers();

                // 1) Try reading serial @1 quickly (robust, but don't block forever)
                int serial = 0;
                var t0 = DateTime.UtcNow;
                while ((DateTime.UtcNow - t0).TotalMilliseconds < 900 && serial <= 0)
                {
                    token.ThrowIfCancellationRequested();
                    try { serial = _serial.ReadSerialFast(1, PollFastMs, token); } catch { /* retry */ }
                    if (serial > 0) break;
                    await Task.Delay(80, token);
                }
                if (serial <= 0) log("Hinweis: Seriennummer @1 nicht sicher lesbar – fahre fort.");

                // 2) Program address 1 -> new
                byte newAdr = (byte)sensor.Address;
                _serial.SetAddress(1, newAdr, token);
                log($"Adresse programmiert: 1 → {newAdr}");

                // 3) Buzzer handling (clear bit 9) BEFORE reboot if requested
                if (string.Equals(sensor.Buzzer, "Disable", StringComparison.OrdinalIgnoreCase))
                {
                    ushort v = _serial.ReadRegisterFast(1, 255, PollFastMs, token);
                    ushort mask = (ushort)(1 << 9);
                    ushort newVal = (ushort)(v & ~mask);
                    if (newVal != v)
                    {
                        await _serial.WriteSingleAsync(255, newVal, 1);
                        log("Buzzer disabled (bit9 cleared).");
                    }
                }

                token.ThrowIfCancellationRequested();

                // 4) Reboot
                _serial.SoftRestart(1, token);
                log("Reboot @1 ausgelöst …");
                await Task.Delay(RebootWaitMs, token);

                // 5) Verify alive @ new address (soft OK if not strictly confirmed)
                var tVerify = DateTime.UtcNow;
                bool moved = false;
                while ((DateTime.UtcNow - tVerify).TotalMilliseconds < PostMoveTimeoutMs)
                {
                    token.ThrowIfCancellationRequested();
                    if (_serial.CheckFast(newAdr, PollFastMs, token)) { moved = true; break; }
                    await Task.Delay(PollDelayMs, token);
                }

                if (!moved)
                {
                    log("Warnung: Gerät antwortet nicht stabil unter neuer Adresse – vermutlich bereits gesetzt. Markiere als OK.");
                    // Try one last serial read @ new address (longer)
                    if (serial <= 0)
                    {
                        try { serial = _serial.ReadSerialFast(newAdr, 800, token); } catch { }
                    }
                    return (true, serial > 0 ? serial : (int?)null);
                }

                // Optional: read a quick register for debug
                try
                {
                    var val = _serial.ReadRegisterFast(newAdr, 2, PollFastMs, token);
                    log($"Gerät antwortet @ {newAdr}. Reg2={val}");
                }
                catch { /* ignore */ }

                // Fallback serial read @ new address if needed
                if (serial <= 0)
                {
                    try { serial = _serial.ReadSerialFast(newAdr, 1200, token); } catch { }
                }

                return (true, serial > 0 ? serial : (int?)null);
            }
            catch (OperationCanceledException)
            {
                log("Abgebrochen.");
                throw;
            }
            catch (Exception ex)
            {
                log($"Config-Fehler: {ex.Message}");
                return (false, null);
            }
        }
    }
}
