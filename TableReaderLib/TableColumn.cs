using System;
using System.Collections.Generic;
using System.Text;

namespace TableReaderLib
{
    public struct TableColumn
    {
        /// <summary>
        /// Название столбца
        /// </summary>
        public string Name
        {
            get;
            set;
        }
        /// <summary>
        /// Тип данных в столбце
        /// </summary>
        public Type Type
        {
            get;
            set;
        }

        /// <summary>
        /// Номер столбца в источнике (от 0)
        /// </summary>
        public int IndexInSource
        {
            get;
            set;
        }
    }
}
