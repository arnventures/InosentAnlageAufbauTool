// File: Converters/WidthToBoolConverter.cs
using System;
using System.Globalization;
using System.Windows.Data;

namespace InosentAnlageAufbauTool.Converters
{
    /// <summary>
    /// Liefert true, wenn die Breite >= Threshold ist.
    /// </summary>
    public class WidthToBoolConverter : IValueConverter
    {
        public double Threshold { get; set; } = 1400;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is double w && w >= Threshold;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
