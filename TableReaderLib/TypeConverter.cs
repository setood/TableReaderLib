using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableReaderLib
{
    public static class TypeConverter
    {
        public static T From<T>(object value)
        {
            //if (CellsValues[index] == null)
            //    return default(T);
            //Type unNullableType = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
            //return (T)Convert.ChangeType(CellsValues[index], unNullableType);

            if (value == null)
                return default(T);
            Type unNullableType = typeof(T);  //Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);

            var converter = TypeDescriptor.GetConverter(unNullableType);
            return converter.CanConvertFrom(value.GetType()) ?
                (T)converter.ConvertFrom(value) :
                throw new InvalidCastException("Can't format " + value.ToString() + "   To " + typeof(T).ToString());
        }
    }
}
