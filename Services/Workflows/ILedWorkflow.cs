using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InosentAnlageAufbauTool.Models;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    public sealed record LedProgress(Led Led, string Status, string? Note);

    public interface ILedWorkflow
    {
        /// <summary>
        /// Programmiert die �bergebenen LEDs sequentiell (nur Adressen > 185).
        /// </summary>
        Task ConfigureAsync(
            IList<Led> lights,
            Action<string> log,
            IProgress<LedProgress> progress,
            CancellationToken token);
    }
}
