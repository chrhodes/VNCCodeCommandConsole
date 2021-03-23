using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Data;
using System.Xml.Linq;

namespace VNCCodeCommandConsole.Presentation.Converters
{
    public class ProjectConverter : IValueConverter
    {
            public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value != null)
            {
                return value;
                //XElement e = (XElement)value;
                //string result = e.Attribute("FileName").Value;
                //if (result == "")
                //{
                //    result = e.Attribute("Name").Value;
                //}
                //return result;
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
            if (value != null)
            {
                ObservableCollection<XElement> collection = new ObservableCollection<XElement>();

                foreach (var item in (value as List<Object>))
                {
                    if (item as string != "")
                    {
                        collection.Add(XElement.Parse(item.ToString()));
                    }
                }

                //ObservableCollection<XElement> collection = value as ObservableCollection<XElement>;

                //return (System.Collections.Generic.List<Object>)value;
                //return "ConvertBack value is not null";
                //List<object> collection = value as List<object>;
                //if (collection is null)
                //    return "no selection";

                //if (collection.Contains("Su") && collection.Contains("Sa"))
                //    return "Weekends";
                //else return "Week days";
                return collection;
            }
            else
            {
                return "ConvertBack value is null";
            }
        }
        
    }
}