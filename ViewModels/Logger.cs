﻿using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows.Controls;
using Thesis.Models;

namespace Thesis.ViewModels
{
    public static class Logger
    {
        private static ObservableCollection<LogItem> _log;
        private static ListView _logView;
        private static Action _selectLogTab;

        public static void Instantiate(ListView logView, Action selectLogTab)
        {
            _log = new ObservableCollection<LogItem>();
            Logger._logView = logView;
            logView.ItemsSource = _log;
            Logger._selectLogTab = selectLogTab;
        }

        public static void Clear()
        {
            _log.Clear();
        }

        public static LogItem Log(LogItemType type, string message, bool useStopwatch = false)
        {
            if (type == LogItemType.Error) _selectLogTab();

            var logItem = new LogItem(type, message, useStopwatch);
            _log.Add(logItem);
            _logView.SelectedIndex = _logView.Items.Count - 1;
            _logView.ScrollIntoView(_logView.SelectedItem);

            return logItem;
        }
    }
}