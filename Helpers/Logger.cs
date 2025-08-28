using System;  // Added for Console

namespace InosentAnlageAufbauTool.Helpers
{
    public interface ILogger
    {
        void Log(string message);
    }

    public class Logger : ILogger
    {
        public void Log(string message)
        {
            Console.WriteLine(message);  // Now resolves
            // File.AppendAllText("log.txt", message + "\n");
        }
    }
}