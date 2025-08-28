using System;
using System.Threading;
using System.Threading.Tasks;

namespace InosentAnlageAufbauTool.Services.Workflows
{
    public interface ILedWorkflow
    {
        /// <summary>
        /// Programs LEDs using Excel (sheet LIGHTIMOK). Only addresses > 185. Timeout (reg 3) gets 0 or 180.
        /// </summary>
        Task ConfigureFromExcelAsync(string excelPath, Action<string> log, CancellationToken token);
    }
}
