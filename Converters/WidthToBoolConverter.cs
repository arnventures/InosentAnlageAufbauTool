// WidthToBoolConverter.cs
using System;
using System.Globalization;
using System.Windows.Data;

namespace InosentAnlageAufbauTool.Converters
{
    // true, wenn Breite >= Threshold
    public class WidthToBoolConverter : IValueConverter
    {
        public double Threshold { get; set; } = 1400;
        public object Convert(object value, Type t, object p, CultureInfo c) => value is double w && w >= Threshold;
        public object ConvertBack(object v, Type t, object p, CultureInfo c) => throw new NotImplementedException();
    }
}
