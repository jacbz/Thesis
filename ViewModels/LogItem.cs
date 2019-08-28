using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace Thesis.ViewModels
{
    public class LogItem : INotifyPropertyChanged
    {
        private string _message;
        private DateTime _time;
        private readonly Stopwatch _stopwatch;

        public LogItem(LogItemType type, string message, bool useStopwatch)
        {
            Type = type;
            Message = message;
            Time = DateTime.Now;
            if (useStopwatch)
                _stopwatch = Stopwatch.StartNew();
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
            var action = new Action(() =>
            {
                if (_stopwatch == null)
                    return;
                Message += $" (elapsed {_stopwatch.ElapsedMilliseconds} ms)";
                Time = DateTime.Now;

                _stopwatch.Stop();
            });

            if (Thread.CurrentThread == Application.Current.Dispatcher.Thread)
                action();
            else
                Application.Current.Dispatcher.Invoke(action, DispatcherPriority.Background);
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