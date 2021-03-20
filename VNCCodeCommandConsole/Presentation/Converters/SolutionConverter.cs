using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Xml.Linq;

namespace VNCCodeCommandConsole.Presentation.Converters
{
    public class SolutionConverter : IValueConverter
    {
            public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value != null)
            {
                XElement e = (XElement)value;
                return e.Attribute("FileName").Value;
                //return "Convert value is not null";
                //List<object> collection = value as List<object>;
                //if (collection.Count == 0)
                //    return "no selection";

                //if (collection.Contains("Su") && collection.Contains("Sa"))
                //    return "Weekends";
                //else return "Week days";
            }
            else
            {
                return "<none selected>";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value;
            //if (value != null)
            //{
            //    return "ConvertBack value is not null";
            //    //List<object> collection = value as List<object>;
            //    //if (collection.Count == 0)
            //    //    return "no selection";

            //    //if (collection.Contains("Su") && collection.Contains("Sa"))
            //    //    return "Weekends";
            //    //else return "Week days";
            //}
            //else return "ConvertBack value is null";
        }
    }
}