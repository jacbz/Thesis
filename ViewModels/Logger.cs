using System.Collections.ObjectModel;
using System.Windows.Controls;

namespace Thesis.ViewModels
{
    public static class Logger
    {
        private static ObservableCollection<LogItem> _log;
        private static ListView _logView;

        public static void Instantiate(ListView logView)
        {
            _log = new ObservableCollection<LogItem>();
            Logger._logView = logView;
            logView.ItemsSource = _log;
        }

        public static LogItem Log(LogItemType type, string message)
        {
            var logItem = new LogItem(type, message);
            _log.Add(logItem);
            _logView.SelectedIndex = _logView.Items.Count - 1;
            _logView.ScrollIntoView(_logView.SelectedItem);
            return logItem;
        }
    }
}