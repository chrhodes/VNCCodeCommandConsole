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
                ObservableCollection<XElement> collection = (ObservableCollection<XElement>)value;

                string result = "";

                foreach (var item in collection)
                {
                    string name = item.Attribute("FileName").Value;

                    result += result == "" ? name : $"; {name}";
                }

                return result;
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
                        try
                        {
                            collection.Add(XElement.Parse(item.ToString()));
                        }
                        catch (System.Xml.XmlException ex)
                        {
                            // This happens after we have picked something
                            // and are trying to pick something else
                            // or clear the selection.
                        }

                    }
                }

                return collection;
            }
            else
            {
                return "ConvertBack value is null";
            }
        }
    }
}