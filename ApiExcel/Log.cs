using Microsoft.Extensions.Logging;

namespace ApiExcel
{
    public class Log
    {
        private static Log _instance;
        private static readonly object _lock = new object();

        
        private Log() { }

        
        public static Log Instance
        {
            get
            {
                lock (_lock)
                {
                    
                    return _instance ??= new Log();
                }
            }
        }

        // Donde registramos los errores
        public void LogError(string message)
        {
            
            var logPath = Path.Combine(@"C:\Logs", $"{DateTime.Now:yyyy-MM-dd}.log");

            using (var writer = new StreamWriter(logPath, true))
            {
                //Fecha con el mensaje de error
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }
    }
}
