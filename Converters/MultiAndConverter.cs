using System;
using System.Globalization;
using System.Windows.Data;

namespace InosentAnlageAufbauTool.Converters
{
    /// <summary>
    /// true, wenn ALLE Eingaben "truthy" sind (bool == true oder nicht-leerer String).
    /// </summary>
    public class MultiAndConverter : IMultiValueConverter
    {
        public static MultiAndConverter Instance { get; } = new MultiAndConverter();

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            foreach (var v in values)
            {
                if (v is bool b) { if (!b) return false; }
                else if (v is string s) { if (string.IsNullOrWhiteSpace(s)) return false; }
                else if (v == null) return false;
            }
            return true;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }
}
