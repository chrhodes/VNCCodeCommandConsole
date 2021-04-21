using System;
using System.Windows.Data;
using System.Xml.Linq;

namespace VNCCodeCommandConsole.Presentation.Converters
{
    public class BranchConverter : IValueConverter
    {
            public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value != null)
            {
                XElement e = (XElement)value;
                return $"{e.Attribute("Repository").Value} - {e.Attribute("Name").Value}";
            }
            else
            {
                return "<none selected>";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value != null)
            {
                return value;
            }
            else
            {
                return "ConvertBack value is null";
            }
        }
    }
}