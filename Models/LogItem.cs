using System;

namespace Thesis
{
    public class LogItem
    {
        public LogItemType Type { get; }
        public string Message { get; }
        public DateTime Time { get; }

        public LogItem(LogItemType type, string message)
        {
            Type = type;
            Message = message;
            Time = DateTime.Now;
        }
    }

    public enum LogItemType
    {
        Info, Warning, Success, Error
    }
}
