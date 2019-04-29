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
            return From<T>(value, null);
        }

        internal static T From<T>(object value, ITypeAdapter typeAdapter)
        {
            if (value == null)
                return default(T);
            Type unNullableType = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
            if (typeAdapter != null)
                value = typeAdapter.Clean(value, unNullableType);
            var valueType = value.GetType();
            if (valueType == unNullableType)
                return (T)value;
            var converter = TypeDescriptor.GetConverter(unNullableType);
            return converter.CanConvertFrom(valueType) ?
                (T)converter.ConvertFrom(value) :
                throw new InvalidCastException($"Can't format {value.ToString()} (Type is {value.GetType()})  To {typeof(T).ToString()}");

        }
    }
}
