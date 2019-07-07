using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Thesis.ViewModels
{
    public static class Logger
    {
        private static ObservableCollection<LogItem> log;
        private static ListView logView;

        public static void Instantiate(ListView logView)
        {
            log = new ObservableCollection<LogItem>();
            Logger.logView = logView;
            logView.ItemsSource = log;
        }

        public static void Log(LogItemType type, string message)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                log.Add(new LogItem(type, message));
                logView.SelectedIndex = logView.Items.Count - 1;
                logView.ScrollIntoView(logView.SelectedItem);
            });
        }
    }
}
