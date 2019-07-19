using OneNoteToFiles.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SongHelper.Services
{
    public struct LogItem
    {
        public string Message { get; set; }
        public string PageId { get; set; }
        public string ContentObjectId { get; set; }

        public override string ToString()
        {
            return Message;
        }

        public static implicit operator string(LogItem item)
        {
            return item.Message;
        }

        public static implicit operator LogItem(string message)
        {
            return new LogItem() { Message = message };
        }
    }

    public static class Logger
    {
        public enum Severity
        {
            Info,
            Warning,
            Error
        }

        public static bool ErrorWasLogged = false;
        public static bool WarningWasLogged = false;
        private static int _level = 0;
        private static string _logFilePath;
        private static FileStream _fileStream = null;
        private static StreamWriter _streamWriter = null;

        public static List<LogItem> Errors { get; set; }
        public static List<LogItem> Warnings { get; set; }

        public static string LogFilePath
        {
            get
            {
                return _logFilePath;
            }
        }

        private static string _errorText = "ОШИБКА: ";

        public static void MoveLevel(int levelDiv)
        {
            _level += levelDiv;
        }

        private static bool _isInitialized = false;

        public static void Init(string systemName)
        {
            if (!_isInitialized)
            {
                string directoryPath = Path.Combine(
                                            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SongHelper"),
                                            "Logs");

                if (!Directory.Exists(directoryPath))
                    Directory.CreateDirectory(directoryPath);

                _logFilePath = Path.Combine(directoryPath, systemName + ".txt");
                try
                {
                    _fileStream = new FileStream(_logFilePath, FileMode.Create);
                }
                catch (IOException)
                {
                    _logFilePath = Path.Combine(directoryPath, string.Format("{0}_{1}.txt", systemName, Guid.NewGuid()));
                    _fileStream = new FileStream(_logFilePath, FileMode.Create);
                }

                _streamWriter = new StreamWriter(_fileStream, Encoding.UTF8);

                Errors = new List<LogItem>();
                Warnings = new List<LogItem>();

                _isInitialized = true;
            }
        }


        public static void Done()
        {
            if (_isInitialized)
            {
                bool needToDelete = false;

                if (_fileStream != null)
                {
                    if (_fileStream.Length == 0)
                        needToDelete = true;
                    _fileStream.Close();
                }


                _isInitialized = false;

                if (needToDelete)
                {
                    try
                    {
                        File.Delete(_logFilePath);
                    }
                    catch { }
                }
                //_lb = null;
            }
        }

        public static void LogMessageEx(string message, bool leveled, bool newLine,
            bool writeDateTime = true, bool silient = false, Severity severity = Severity.Info, string pageId = null, string contentObjectId = null)
        {
            LogMessageToFileAndConsole(false, string.Empty, null, writeDateTime, silient, severity, pageId, contentObjectId);

            if (leveled)
                for (int i = 0; i < _level; i++)
                    LogMessageToFileAndConsole(false, "  ", null, false, silient, severity, pageId, contentObjectId);

            LogMessageToFileAndConsole(newLine, message, null, false, silient, severity, pageId, contentObjectId);
        }
        
        private static void LogMessageToFileAndConsole(bool newLine, string message, string messageEx, bool writeDateTime, bool silient, Severity severity, string pageId, string contentObjectId)
        {
            if (!_isInitialized)
            {
                try
                {
                    Init(System.Reflection.Assembly.GetEntryAssembly().GetName().Name);
                }
                catch { }
            }

            if (string.IsNullOrEmpty(messageEx))
                messageEx = message;

            if (writeDateTime)
                messageEx = string.Format("{0}: {1}", DateTime.Now, messageEx);


            if (newLine)
            {
                Console.WriteLine(message);

                TryToWriteToFile(messageEx, newLine);
            }
            else
            {
                Console.Write(message);

                TryToWriteToFile(messageEx, newLine);
            }

            try
            {
                if (_streamWriter != null && _streamWriter.BaseStream != null)
                    _streamWriter.Flush();
            }
            catch { }

            if (severity == Severity.Error)
            {
                if (Errors != null && !string.IsNullOrEmpty(message) && !string.IsNullOrEmpty(message.Trim()))
                {
                    Errors.Add(new LogItem()
                    {
                        Message = message,
                        PageId = pageId,
                        ContentObjectId = contentObjectId
                    });
                }
            }
            else if (severity == Severity.Warning)
            {
                if (Warnings != null && !string.IsNullOrEmpty(message) && !string.IsNullOrEmpty(message.Trim()))
                {
                    Warnings.Add(new LogItem()
                    {
                        Message = message,
                        PageId = pageId,
                        ContentObjectId = contentObjectId
                    });
                }
            }
        }

        private static void TryToWriteToFile(string message, bool newLine)
        {
            if (_streamWriter != null && _streamWriter.BaseStream != null)
            {
                if (newLine)
                    _streamWriter.WriteLine(message);
                else
                    _streamWriter.Write(message);
            }
        }


        public static void LogWarning(string message, params object[] args)
        {
            LogMessageEx("Warning: " + FormatString(message, args), true, true, true, false, Severity.Warning);
            WarningWasLogged = true;
        }

        public static void LogWarning(string pageId, string contentObjectId, string message, params object[] args)
        {
            LogMessageEx("Warning: " + FormatString(message, args), true, true, true, false, Severity.Warning, pageId, contentObjectId);
            WarningWasLogged = true;
        }

        /// <summary>
        /// Log only to log
        /// </summary>
        /// <param name="message"></param>
        /// <param name="args"></param>
        public static void LogMessageSilient(string message, params object[] args)
        {
            LogMessageEx(FormatString(message, args), true, true, true, true);
        }

        public static void LogMessage(string message, params object[] args)
        {
            LogMessageEx(FormatString(message, args), true, true);
        }

        public static void LogError(string message, Exception ex)
        {
            LogMessageToFileAndConsole(true, string.Format("{0}{1} {2}", _errorText, message, OneNoteUtils.ParseErrorAndMakeItMoreUserFriendly(ex.Message)), string.Format("{0} {1}", message, ex.ToString()), true, false, Severity.Error, null, null);
            ErrorWasLogged = true;
        }

        public static void LogError(Exception ex)
        {
            LogError(string.Empty, ex);
        }

        public static void LogError(string message, params object[] args)
        {
            LogMessageToFileAndConsole(true, _errorText + FormatString(message, args), null, true, false, Severity.Error, null, null);
            ErrorWasLogged = true;
        }

        private static string FormatString(string message, params object[] args)
        {
            return args.Count() == 0 ? message : string.Format(message, args);
        }
    }
}
