// MultiOrConverter.cs
using System;
using System.Globalization;
using System.Windows.Data;

namespace InosentAnlageAufbauTool.Converters
{
    // true, wenn irgendein Input "truthy" ist (bool==true ODER nicht-leerer String)
    public class MultiOrConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type t, object p, CultureInfo c)
        {
            foreach (var v in values)
            {
                if (v is bool b && b) return true;
                if (v is string s && !string.IsNullOrWhiteSpace(s)) return true;
            }
            return false;
        }
        public object[] ConvertBack(object v, Type[] ts, object p, CultureInfo c) => throw new NotImplementedException();
    }
}
