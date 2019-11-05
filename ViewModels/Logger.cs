// Thesis - An Excel to code converter
// Copyright (C) 2019 Jacob Zhang
// 
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

using System;
using System.Collections.ObjectModel;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

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
            _logView = logView;
            logView.ItemsSource = _log;
            _selectLogTab = selectLogTab;
        }

        public static void Clear()
        {
            _log.Clear();
        }

        public static LogItem Log(LogItemType type, string message, bool useStopwatch = false)
        {
            var createLogItem = new Func<LogItem>(() =>
            {
                if (type == LogItemType.Error) _selectLogTab();

                var logItem = new LogItem(type, message, useStopwatch);
                _log.Add(logItem);
                _logView.SelectedIndex = _logView.Items.Count - 1;
                _logView.ScrollIntoView(_logView.SelectedItem);

                return logItem;
            });

            if (Thread.CurrentThread == Application.Current.Dispatcher.Thread)
                return createLogItem();
            return Application.Current.Dispatcher.Invoke(createLogItem, DispatcherPriority.Background);
        }
    }
}