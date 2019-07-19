namespace OneNoteToFiles.Contracts
{
    public interface ICustomLogger
    {
        bool AbortedByUser { get; }

        void LogMessage(string message, params object[] args);
        void LogWarning(string message, params object[] args);
        void LogException(string message, params object[] args);
    }
}
