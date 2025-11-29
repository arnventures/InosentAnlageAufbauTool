// File: Converters/MultiOrConverter.cs
using System;
using System.Globalization;
using System.Windows.Data;

namespace InosentAnlageAufbauTool.Converters
{
    /// <summary>
    /// true, wenn irgendein Input "truthy" ist (bool == true oder nicht-leerer String).
    /// </summary>
    public class MultiOrConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null)
                return false;

            foreach (var v in values)
            {
                if (v is bool b && b)
                    return true;

                if (v is string s && !string.IsNullOrWhiteSpace(s))
                    return true;
            }

            return false;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
