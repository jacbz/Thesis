using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Thesis
{
    public class LogItem : INotifyPropertyChanged
    {
        private string message;
        private DateTime time;

        public LogItem(LogItemType type, string message)
        {
            Type = type;
            Message = message;
            Time = DateTime.Now;
        }

        public LogItemType Type { get; }

        public string Message
        {
            get => message;
            set
            {
                message = value;
                OnPropertyChanged();
            }
        }

        public DateTime Time
        {
            get => time;
            set
            {
                time = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

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

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public enum LogItemType
    {
        Info,
        Warning,
        Success,
        Error
    }
}