using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InosentAnlageAufbauTool.Models;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    /// <summary>
    /// UI-Progress-Payload fuer Status-Updates je Sensor.
    /// </summary>
    /// <param name="Sensor">Das Sensor-Objekt (fuer direkte Bindung/Update).</param>
    /// <param name="Status">"Active", "OK", "Fail", "Skipped".</param>
    /// <param name="Serial">Gelesene Seriennummer (falls vorhanden).</param>
    /// <param name="Note">Optionaler Hinweis/Fehlermeldungstext.</param>
    public sealed record SensorProgress(Sensor Sensor, string Status, int? Serial, string? Note);

    public sealed record SensorRunResult(int ExcelRowIndex, int Serial);

    public interface ISensorWorkflow
    {
        /// <summary>
        /// Programmiert alle uebergebenen Sensoren sequenziell:
        /// - wartet auf neues Geraet @1 (entprellt)
        /// - liest Seriennummer @1 (robust, schnell)
        /// - setzt Modbus-Adresse, optional Buzzer-Flag, Reboot
        /// - verifiziert (soft OK, um letzte-Fehlschlag-Falle zu vermeiden)
        /// - sammelt SNs fuer Batch-Write (durch Aufrufer gespeichert, auch bei Abbruch).
        /// </summary>
        /// <param name="sensors">Liste der Ziel-Sensoren (enthaelt Address/Buzzer/Index).</param>
        /// <param name="skipRequested">Delegate, der bei <c>true</c> den aktuellen Schritt ueberspringt.</param>
        /// <param name="progress">Progress fuer UI-Status je Sensor.</param>
        /// <param name="log">Log-Callback fuer Protokollausgaben.</param>
        /// <param name="token">CancellationToken (Stop-Button).</param>
        Task<IReadOnlyList<SensorRunResult>> RunAsync(
            IList<Sensor> sensors,
            Func<bool> skipRequested,
            IProgress<SensorProgress> progress,
            Action<string> log,
            CancellationToken token);
    }
}
