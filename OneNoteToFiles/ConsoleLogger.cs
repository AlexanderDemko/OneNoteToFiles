using OneNoteToFiles.Contracts;
using System;

namespace OneNoteToFiles
{
    public class ConsoleLogger: ICustomLogger
    {
        public void LogMessage(string message, params object[] args)
        {
            Console.WriteLine(message, args);
        }

        public void LogWarning(string message, params object[] args)
        {
            Console.WriteLine(message, args);
        }

        public void LogException(string message, params object[] args)
        {
            Console.WriteLine(message, args);
        }

        public bool AbortedByUser
        {
            get { return false; }
        }
    }
}
