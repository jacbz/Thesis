using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Thesis
{
    public class LogItem : INotifyPropertyChanged
    {
        public LogItemType Type { get; }
        private string message;
        public string Message
        {
            get => message;
            set
            {
                message = value;
                OnPropertyChanged("Message");
            }
        }
        private DateTime time;
        public DateTime Time
        {
            get => time;
            set
            {
                time = value;
                OnPropertyChanged("Time");
            }
        }

        public void AppendTime(long ms)
        {
            Message += $" (elapsed {ms} ms)";
            Time = DateTime.Now;
        }

        public void Append(string message)
        {
            Message += " " + message;
            Time = DateTime.Now;
        }

        public LogItem(LogItemType type, string message)
        {
            Type = type;
            Message = message;
            Time = DateTime.Now;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    public enum LogItemType
    {
        Info, Warning, Success, Error
    }
}
