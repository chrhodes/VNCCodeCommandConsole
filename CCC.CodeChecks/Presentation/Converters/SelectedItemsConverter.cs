using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;

namespace CCC.CodeChecks.Presentation.Converters
{
    public class SelectedItemsConverter : MarkupExtension, IValueConverter
    {
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }

        object IValueConverter.Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
                return new List<object>((IEnumerable<string>)value);

            return null;
        }
        //object IValueConverter.ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        //{
        //    ObservableCollection<string> result = new ObservableCollection<string>();
        //    var enumerable = (List<object>)value;
        //    if (enumerable != null)
        //        foreach (object item in enumerable)
        //            result.Add((string)item);
        //    return result;
        //}

        object IValueConverter.ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            List<string> result = new List<string>();
            var enumerable = (List<object>)value;
            if (enumerable != null)
                foreach (object item in enumerable)
                    result.Add((string)item);
            return result;
        }
    }
}
