using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableReaderLib
{
    /// <summary>
    /// Помогает приводить специфичные типы данных в стандартные.
    /// </summary>
    public interface ITypeAdapter
    {
        /// <summary>
        /// Format  value  from Tin to Tout type 
        /// </summary>
        /// <returns></returns>
        object Clean(object value, Type typeToConvert);
    }
}
