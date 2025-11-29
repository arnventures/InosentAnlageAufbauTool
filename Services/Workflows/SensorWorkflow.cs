using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InosentAnlageAufbauTool.Helpers;
using InosentAnlageAufbauTool.Models;
using InosentAnlageAufbauTool.Services;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    public sealed class SensorWorkflow : ISensorWorkflow
    {
        private readonly SerialService _serial;
        private readonly ILogger? _logger;

        // Tuned timings (tight for speed, still tolerant)
        private const int PollFastMs = 50;
        private const int PollDelayMs = 50;
        private const int RebootWaitMs = 600;
        private const int VerifyTimeoutMs = 1400;

        public SensorWorkflow(SerialService serial, ILogger? logger = null)
        {
            _serial = serial;
            _logger = logger;
        }

        public async Task<IReadOnlyList<SensorRunResult>> RunAsync(
            IList<Sensor> sensors,
            Func<bool> skipRequested,
            IProgress<SensorProgress> progress,
            Action<string> log,
            CancellationToken token)
        {
            var batchSerials = new List<SensorRunResult>();

            foreach (var sensor in sensors)
            {
                token.ThrowIfCancellationRequested();

                progress.Report(new SensorProgress(sensor, "Active", sensor.Serial, null));

                log($"Warte auf Geraet @1 fuer Sensor {sensor.Index} ... (Skip ueberspringt; Stop beendet)");

                var ready = await WaitForAddr1Async(skipRequested, token);
                if (!ready)
                {
                    if (token.IsCancellationRequested) throw new OperationCanceledException(token);
                    progress.Report(new SensorProgress(sensor, "Skipped", sensor.Serial, "vom Benutzer uebersprungen"));
                    log($"Sensor {sensor.Index} uebersprungen.");
                    continue;
                }

                var (success, serial) = await ConfigureSingleAsync(sensor, log, token);
                if (success && serial.HasValue && serial.Value > 0)
                {
                    batchSerials.Add(new SensorRunResult(sensor.ExcelRow, serial.Value));
                    sensor.Serial = serial.Value;
                }

                var newStatus = success ? "OK" : "Fail";
                progress.Report(new SensorProgress(sensor, newStatus, serial, null));
                log(success
                    ? $"Sensor {sensor.Index} OK (SN={serial?.ToString() ?? "-"})"
                    : $"Sensor {sensor.Index} Fail.");
            }

            return batchSerials;
        }

        private async Task<bool> WaitForAddr1Async(Func<bool> skipRequested, CancellationToken token)
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
                catch { /* transient */ }

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

                // 1) Quick serial read on addr 1
                int serial = TryReadSerialQuick(1, attempts: 6, timeoutMs: PollFastMs + 40, token);

                // 2) Optional buzzer disable before address change
                if (string.Equals(sensor.Buzzer, "Disable", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        ushort current = _serial.ReadRegisterFast(1, 255, PollFastMs, token);
                        ushort mask = (ushort)(1 << 9);
                        ushort newVal = (ushort)(current & ~mask);
                        if (newVal != current)
                        {
                            await _serial.WriteSingleAsync(255, newVal, 1);
                            log("Buzzer disabled (bit9 cleared).");
                        }
                    }
                    catch (Exception ex)
                    {
                        log($"Buzzer disable Hinweis: {ex.Message}");
                    }
                }

                // 3) Program address 1 -> new
                byte newAdr = (byte)sensor.Address;
                _serial.SetAddress(1, newAdr, token);
                log($"Adresse programmiert: 1 -> {newAdr}");

                // 4) Soft reboot to lock in address
                _serial.SoftRestart(1, token);
                await Task.Delay(RebootWaitMs, token);

                // 5) Verify alive @ new address (lenient)
                bool moved = await ConfirmOnAddressAsync(newAdr, token);
                if (!moved)
                {
                    log("Hinweis: Ger�t antwortet nicht stabil unter neuer Adresse - fahre fort.");
                }

                // 6) Fallback serial read @ new address if needed
                if (serial <= 0)
                {
                    serial = TryReadSerialQuick(newAdr, attempts: 5, timeoutMs: 300, token);
                }

                if (serial <= 0)
                    log("Seriennummer nicht lesbar.");
                else
                    log($"Seriennummer = {serial}");

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
                _logger?.Log($"[SensorWorkflow] Fehler: {ex.Message}");
                return (false, null);
            }
        }

        private async Task<bool> ConfirmOnAddressAsync(byte address, CancellationToken token)
        {
            var start = DateTime.UtcNow;
            while ((DateTime.UtcNow - start).TotalMilliseconds < VerifyTimeoutMs)
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    if (_serial.CheckFast(address, PollFastMs, token))
                        return true;
                }
                catch (OperationCanceledException) { throw; }
                catch { /* transient */ }

                try { await Task.Delay(PollDelayMs, token); } catch (OperationCanceledException) { throw; }
            }
            return false;
        }

        private int TryReadSerialQuick(byte unit, int attempts, int timeoutMs, CancellationToken token)
        {
            int serial = 0;
            for (int i = 0; i < attempts && serial <= 0; i++)
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    serial = _serial.ReadSerialFast(unit, timeoutMs, token);
                }
                catch
                {
                    serial = 0;
                }

                if (serial > 0) break;
                if (i < attempts - 1)
                {
                    try { Task.Delay(PollDelayMs, token).GetAwaiter().GetResult(); }
                    catch (OperationCanceledException) { throw; }
                }
            }
            return serial;
        }

    }
}
