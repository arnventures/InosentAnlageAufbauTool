// File: Converters/BoolToBrushConverter.cs
using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace InosentAnlageAufbauTool.Converters
{
    public class BoolToBrushConverter : IValueConverter
    {
        // Nullable, damit keine Nullable-Error-Warnungen entstehen
        public Brush? TrueBrush { get; set; } = Brushes.Green;
        public Brush? FalseBrush { get; set; } = Brushes.Gray;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var brush = (value is bool b && b) ? TrueBrush : FalseBrush;
            return brush ?? Brushes.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
