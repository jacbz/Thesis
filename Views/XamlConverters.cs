using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using Thesis.Models.VertexTypes;
using Thesis.ViewModels;

namespace Thesis.Views
{
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
