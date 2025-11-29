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
    public sealed class SensorWorkflow : ISensorWorkflow
    {
        private readonly SerialService _serial;
        private readonly ILogger? _logger;

        // Tuned timings for quicker but stable loops
        private const int PollIntervalMs = 60;
        private const int StableWindowMs = 180;
        private const int PresenceTimeoutMs = 140;
        private const int SerialReadTimeoutMs = 170;
        private const int SerialReadAttempts = 6;
        private const int SerialReadAttemptsAfterMove = 8;
        private const int AddressHandoverGapMs = 110;
        private const int RebootWaitMs = 450;
        private const int WaitGoneTimeoutMs = 1800;
        private const int WaitAliveTimeoutMs = 1400;

        private const ushort BuzzerRegister = 255;
        private const ushort BuzzerBit = 1 << 9;

        private enum WaitOutcome
        {
            Ready,
            Skip,
            Cancel
        }

        private sealed record SensorOutcome(bool Success, int? Serial, string? Note);

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
                if (token.IsCancellationRequested) break;

                if (!sensor.IsSelected)
                {
                    progress.Report(new SensorProgress(sensor, "Skipped", sensor.Serial, "nicht ausgewaehlt"));
                    log($"Sensor {sensor.Index} uebersprungen (nicht ausgewaehlt).");
                    continue;
                }

                progress.Report(new SensorProgress(sensor, "Active", sensor.Serial, null));
                log($"Sensor {sensor.Index}: warte auf Geraet @1 ... (Skip ueberspringt; Stop beendet)");

                var waitOutcome = await WaitForAddr1Async(skipRequested, token);
                if (waitOutcome == WaitOutcome.Cancel) break;
                if (waitOutcome == WaitOutcome.Skip)
                {
                    progress.Report(new SensorProgress(sensor, "Skipped", sensor.Serial, "vom Benutzer uebersprungen"));
                    log($"Sensor {sensor.Index} uebersprungen.");
                    continue;
                }

                SensorOutcome outcome;
                try
                {
                    outcome = await ConfigureSingleAsync(sensor, log, token);
                }
                catch (OperationCanceledException)
                {
                    progress.Report(new SensorProgress(sensor, "Fail", sensor.Serial, "Abgebrochen"));
                    break;
                }
                if (outcome.Serial.HasValue && outcome.Serial.Value > 0)
                {
                    batchSerials.Add(new SensorRunResult(sensor.ExcelRow, outcome.Serial.Value));
                    sensor.Serial = outcome.Serial.Value;
                }

                var status = outcome.Success ? "OK" : "Fail";
                progress.Report(new SensorProgress(sensor, status, outcome.Serial, outcome.Note));
                log(outcome.Success
                    ? $"Sensor {sensor.Index} OK (Adr {sensor.Address}, SN={outcome.Serial?.ToString() ?? "-"})"
                    : $"Sensor {sensor.Index} Fail: {outcome.Note ?? "Unbekannter Fehler"}");
            }

            return batchSerials;
        }

        private async Task<WaitOutcome> WaitForAddr1Async(Func<bool> skipRequested, CancellationToken token)
        {
            int stableMs = 0;
            while (true)
            {
                if (token.IsCancellationRequested) return WaitOutcome.Cancel;
                if (skipRequested()) return WaitOutcome.Skip;

                _serial.FlushBuffers();

                bool alive = ProbeAlive(1, token, PresenceTimeoutMs);
                if (!alive)
                {
                    alive = ProbeByType(1, token);
                }

                if (alive)
                {
                    stableMs += PollIntervalMs;
                    if (stableMs >= StableWindowMs) return WaitOutcome.Ready;
                }
                else
                {
                    stableMs = 0;
                }

                try { await Task.Delay(PollIntervalMs, token); }
                catch (OperationCanceledException) { return WaitOutcome.Cancel; }
            }
        }

        private bool ProbeAlive(byte unit, CancellationToken token, int timeoutMs)
        {
            try { return _serial.CheckFast(unit, timeoutMs, token); }
            catch (OperationCanceledException) { throw; }
            catch { return false; }
        }

        private bool ProbeByType(byte unit, CancellationToken token)
        {
            try
            {
                var dt = _serial.ReadRegisterFast(unit, SerialService.REG_DEVICE_TYPE, PresenceTimeoutMs, token);
                return dt > 0;
            }
            catch (OperationCanceledException) { throw; }
            catch { return false; }
        }

        private async Task<SensorOutcome> ConfigureSingleAsync(
            Sensor sensor,
            Action<string> log,
            CancellationToken token)
        {
            bool success = false;
            string? note = null;
            int? serial = null;

            try
            {
                _serial.FlushBuffers();

                var sn = await ReadSerialRobustAsync(1, SerialReadAttempts, SerialReadTimeoutMs, token);
                if (sn > 0) serial = sn;

                await EnsureBuzzerAsync(sensor.Buzzer, log, token);

                var targetAddr = (byte)sensor.Address;
                if (IsAlive(targetAddr, token))
                {
                    note = $"Adresse {targetAddr} ist bereits belegt.";
                    return new SensorOutcome(false, serial, note);
                }

                await RtuGapAsync(token);

                _serial.SetAddress(1, targetAddr, token);
                log($"Adresse programmiert: 1 -> {targetAddr}");

                await RtuGapAsync(token);
                _serial.SoftRestart(1, token);
                await Task.Delay(RebootWaitMs, token);

                var gone = await WaitGoneAsync(1, WaitGoneTimeoutMs, token);
                if (!gone) note = "Addr 1 antwortet weiter (Handover).";

                await RtuGapAsync(token);

                var alive = await WaitAliveStableAsync(targetAddr, WaitAliveTimeoutMs, token);
                if (!alive)
                {
                    note = note ?? "Neue Adresse antwortet nicht stabil.";
                    return new SensorOutcome(false, serial, note);
                }

                if (!serial.HasValue || serial.Value <= 0)
                {
                    sn = await ReadSerialRobustAsync(targetAddr, SerialReadAttemptsAfterMove, 260, token);
                    if (sn > 0) serial = sn;
                }

                if (serial == null) note = "Seriennummer nicht lesbar.";
                success = true;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                note = ex.Message;
                _logger?.Log($"[SensorWorkflow] Fehler {sensor.Address}: {ex.Message}");
                success = false;
            }

            return new SensorOutcome(success, serial, note);
        }

        private async Task EnsureBuzzerAsync(string? buzzerMode, Action<string> log, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(buzzerMode)) return;

            bool disable = string.Equals(buzzerMode, "Disable", StringComparison.OrdinalIgnoreCase);
            bool enable = string.Equals(buzzerMode, "Enable", StringComparison.OrdinalIgnoreCase);

            if (!disable && !enable) return;

            try
            {
                ushort current = _serial.ReadRegisterFast(1, BuzzerRegister, PresenceTimeoutMs, token);
                bool currentlyDisabled = (current & BuzzerBit) == 0;
                ushort desired = disable ? (ushort)(current & ~BuzzerBit) : (ushort)(current | BuzzerBit);

                if (desired != current)
                {
                    await _serial.WriteSingleAsync(BuzzerRegister, desired, 1);
                    log(disable ? "Buzzer disabled (bit9 cleared)." : "Buzzer enabled (bit9 set).");
                }
                else
                {
                    log(disable == currentlyDisabled ? "Buzzer bereits korrekt." : "Buzzer unveraendert.");
                }
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                log($"Buzzer Hinweis: {ex.Message}");
            }
        }

        private async Task<bool> WaitGoneAsync(byte address, int timeoutMs, CancellationToken token)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (!IsAlive(address, token))
                    return true;

                try { await Task.Delay(PollIntervalMs, token); }
                catch (OperationCanceledException) { return false; }
            }
            return false;
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

        private async Task<int> ReadSerialRobustAsync(byte address, int attempts, int timeoutMs, CancellationToken token)
        {
            for (int i = 0; i < attempts; i++)
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    int serial = _serial.ReadSerialFast(address, timeoutMs, token);
                    if (serial > 0) return serial;
                }
                catch (OperationCanceledException) { throw; }
                catch { /* retry */ }

                if (i < attempts - 1)
                {
                    try { await Task.Delay(PollIntervalMs, token); }
                    catch (OperationCanceledException) { throw; }
                }
            }
            return 0;
        }

        private static Task RtuGapAsync(CancellationToken token)
            => Task.Delay(AddressHandoverGapMs, token);
    }
}
