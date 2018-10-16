using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace SharePointOnline_MigrationTool
{
    /// <summary>
    /// Convert a full path to a specific image type of a library
    /// </summary>
    [ValueConversion(typeof(string), typeof(BitmapImage))]
    public class HeaderToImageConverter : IValueConverter
    {
        public static HeaderToImageConverter Instance = new HeaderToImageConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            //Get the TreeviewItem Tag property
            var tag = (string)value;

            //By default, we set UFO
            var image = "Default.png";

            //Icone Distribution Logic
            if (tag.StartsWith("https://")) image = "SPSite.png";
            if (tag == "100") image = "List.png" ;
            if (tag == "101") image = "Lib.png";


            return new BitmapImage(new Uri($"pack://application:,,,/img/{image}"));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
