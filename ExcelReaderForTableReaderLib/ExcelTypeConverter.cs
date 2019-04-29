using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderForTableReaderLib
{
    internal class ExcelTypeConverter : TableReaderLib.ITypeAdapter
    {

        static ExcelTypeConverter singletonInstance;

        internal static ExcelTypeConverter Converter
        {
            get
            {
                if (singletonInstance == null)
                    singletonInstance = new ExcelTypeConverter();

                return singletonInstance;
            }
        }
        private ExcelTypeConverter() { }

        public object Clean(object value, Type typeToConvert)
        {
            if (typeToConvert == typeof(string))
                return value.ToString();
            if (typeToConvert == typeof(int))
                return Convert.ToInt32(value);
            if (typeToConvert == typeof(DateTime))
               return DateTime.FromOADate((double)value);
            if (typeToConvert == typeof(TimeSpan))
            {
                var datetime = DateTime.FromOADate((double)value);
                return datetime.TimeOfDay;
            }

            return value;
        }
    }
}
