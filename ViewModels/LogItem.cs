using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace Thesis.ViewModels
{
    public class LogItem : INotifyPropertyChanged
    {
        private string _message;
        private DateTime _time;
        private Stopwatch _stopwatch;

        public LogItem(LogItemType type, string message, Stopwatch stopwatch)
        {
            Type = type;
            Message = message;
            Time = DateTime.Now;
            _stopwatch = stopwatch;
        }

        public LogItemType Type { get; }

        public string Message
        {
            get => _message;
            set
            {
                _message = value;
                OnPropertyChanged();
            }
        }

        public DateTime Time
        {
            get => _time;
            set
            {
                _time = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void AppendElapsedTime()
        {
            Message += $" (elapsed {_stopwatch.ElapsedMilliseconds} ms)";
            Time = DateTime.Now;

            _stopwatch.Stop();
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