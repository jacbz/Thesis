using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Thesis
{
    public class LogItem
    {
        public LogItemType Type { get; set; }
        public string Message { get; set; }
        public DateTime Time { get; set; }

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
