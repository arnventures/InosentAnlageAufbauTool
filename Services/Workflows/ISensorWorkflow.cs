using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InosentAnlageAufbauTool.Models;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    /// <summary>
    /// UI-Progress-Payload f�r Status-Updates je Sensor.
    /// </summary>
    /// <param name="Sensor">Das Sensor-Objekt (f�r direkte Bindung/Update).</param>
    /// <param name="Status">"Active", "OK", "Fail", "Skipped".</param>
    /// <param name="Serial">Gelesene Seriennummer (falls vorhanden).</param>
    /// <param name="Note">Optionaler Hinweis/Fehlermeldungstext.</param>
    public sealed record SensorProgress(Sensor Sensor, string Status, int? Serial, string? Note);

    public sealed record SensorRunResult(int ExcelRowIndex, int Serial);

    public interface ISensorWorkflow
    {
        /// <summary>
        /// Programmiert alle �bergebenen Sensoren sequenziell:
        /// - wartet auf neues Ger�t @1 (entprellt)
        /// - liest Seriennummer @1 (robust, schnell)
        /// - setzt Modbus-Adresse, optional Buzzer-Flag, Reboot
        /// - verifiziert (soft OK, um letzte-Fehlschlag-Falle zu vermeiden)
        /// - sammelt SNs f�r Batch-Write (durch Aufrufer gespeichert).
        /// </summary>
        /// <param name="sensors">Liste der Ziel-Sensoren (enth�lt Address/Buzzer/Index).</param>
        /// <param name="skipRequested">Delegate, der bei <c>true</c> den aktuellen Schritt �berspringt.</param>
        /// <param name="progress">Progress f�r UI-Status je Sensor.</param>
        /// <param name="log">Log-Callback f�r Protokollausgaben.</param>
        /// <param name="token">CancellationToken (Stop-Button).</param>
        Task<IReadOnlyList<SensorRunResult>> RunAsync(
            IList<Sensor> sensors,
            Func<bool> skipRequested,
            IProgress<SensorProgress> progress,
            Action<string> log,
            CancellationToken token);
    }
}
