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
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using Thesis.Models;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Views
{
    public class NameConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Name name)
                return name.ToString();
            return "";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            Random random = new Random();
            // append random string as address to use if it's a duplicate
            if (value is string name)
                return new Name(name, new string(Enumerable.Repeat("abcdefghijklmnopqrstuvwxyz", 5)
                    .Select(s => s[random.Next(s.Length)]).ToArray()));
            return new Name("", "");
        }
    }
    public class TypeOfConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value?.GetType().Name;
        }

        public object ConvertBack( object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Not needed
            throw new Exception();
        }
    }

    [ValueConversion(typeof(System.Drawing.Color), typeof(SolidColorBrush))]
    public class ColorToSolidBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            System.Drawing.Color color = (System.Drawing.Color)value;
            return new SolidColorBrush(color.ToMColor());
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Not needed
            throw new Exception();
        }
    }

    public class CellVertexVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is CellVertex ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Not needed
            throw new Exception();
        }
    }
}
