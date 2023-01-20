using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows;

namespace EmployerPerformanceCalculator
{
    [ValueConversion(typeof(KeyValuePair<string, string>), typeof(string))]
    internal class PerformanceItemConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            KeyValuePair<string, string> date = (KeyValuePair<string, string>)value;

            return string.Format("{0}<->{1}", date.Key, date.Value);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string strValue = value as string;
            string[] temp = strValue.Split("<->");
            KeyValuePair<string, string> keyValuePairTemp = new KeyValuePair<string, string>(temp[0], temp[1]);
            return keyValuePairTemp;
        }
    }
}
