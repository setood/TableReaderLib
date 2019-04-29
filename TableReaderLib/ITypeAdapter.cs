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
        /// <typeparam name="Tin"></typeparam>
        /// <typeparam name="Tout"></typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        T Clean<T>(object value);
    }
}
