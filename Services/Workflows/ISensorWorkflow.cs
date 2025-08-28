using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InosentAnlageAufbauTool.Models;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    /// <summary>
    /// UI-Progress-Payload für Status-Updates je Sensor.
    /// </summary>
    /// <param name="Sensor">Das Sensor-Objekt (für direkte Bindung/Update).</param>
    /// <param name="Status">"Active", "OK", "Fail", "Skipped" …</param>
    /// <param name="Serial">Gelesene Seriennummer (falls vorhanden).</param>
    /// <param name="Note">Optionaler Hinweis/Fehlermeldungstext.</param>
    public sealed record SensorProgress(Sensor Sensor, string Status, int? Serial, string? Note);

    public interface ISensorWorkflow
    {
        /// <summary>
        /// Programmiert alle übergebenen Sensoren sequenziell:
        /// - wartet auf neues Gerät @1 (entprellt)
        /// - liest Seriennummer @1 (robust, schnell)
        /// - setzt Modbus-Adresse, optional Buzzer-Flag, Reboot
        /// - verifiziert (soft OK, um letzte-Fehlschlag-Falle zu vermeiden)
        /// - sammelt SNs und schreibt sie am Ende in einem Batch in Excel.
        /// </summary>
        /// <param name="sensors">Liste der Ziel-Sensoren (enthält Address/Buzzer/Index).</param>
        /// <param name="excelPath">Pfad zur Excel-Datei (für Batch-Write am Ende).</param>
        /// <param name="skipRequested">Delegate, der bei <c>true</c> den aktuellen Schritt überspringt.</param>
        /// <param name="progress">Progress für UI-Status je Sensor.</param>
        /// <param name="log">Log-Callback für Protokollausgaben.</param>
        /// <param name="token">CancellationToken (Stop-Button).</param>
        Task RunAsync(
            IList<Sensor> sensors,
            string excelPath,
            Func<bool> skipRequested,
            IProgress<SensorProgress> progress,
            Action<string> log,
            CancellationToken token);
    }
}
