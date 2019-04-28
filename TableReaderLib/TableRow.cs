using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.ComponentModel;

namespace TableReaderLib
{
    public sealed class TableRow : IEnumerable<object>
    {
        /// <summary>
        /// Список доступных для чтения столбцов строки
        /// </summary>
        public IReadOnlyList<TableColumn> Columns { get;}
        /// <summary>
        ///данные ячеек
        /// </summary>
        IReadOnlyList<object> CellsValues { get; }
        /// <summary>
        /// Получить значение ячейки из столбца с именем columnName
        /// </summary>
        /// <param name="columnName">Имя столбца из коллекции Columns</param>
        /// <exception cref="ArgumentException">Если имя столбца отсутствует в коллекии Columns</exception>
        /// <returns></returns>
        public object this[string columnName] => GetCellValue<object>(columnName);
        /// <summary>
        /// Получить значение ячейки из столбца column
        /// </summary>
        /// <param name="column">столбец из коллекции Columns</param>
        /// <exception cref="ArgumentException">Если имя столбец отсутствует в коллекии Columns</exception>
        /// <returns></returns>
        public object this[TableColumn column] => GetCellValue<object>(column);
        /// <summary>
        /// Получить значение ячейки из столбца с номером index в  Columns
        /// </summary>
        /// <param name="columnName">Номер столбца из коллекции Columns</param>
        /// <returns></returns>
        public object this[int index] => GetCellValue<object>(index);

        public TableRow(IReadOnlyList<TableColumn> columns, IReadOnlyList<object> cellsValues)
        {
            this.Columns = columns;
            this.CellsValues = cellsValues;
        }

        /// <summary>
        /// Получить значение ячейки из столбца с именем columnName. 
        /// Для приведения значения в тип T используется контрукция: (T)Convert.ChangeType(CellsValues[index], typeof(T));
        /// </summary>
        /// <param name="columnName">Имя столбца из коллекции Columns</param>
        /// <exception cref="ArgumentException">Если имя столбца отсутствует в коллекии Columns</exception>
        /// <returns></returns>
        public T GetCellValue<T>(string columnName)
        {
            if (Columns.Count(c => c.Name == columnName) == 0)
                throw new ArgumentException("Column not in Columns collection");
            var column = Columns.First(c => c.Name == columnName);
            return GetCellValue<T>(column);
        }
        /// <summary>
        /// Получить значение ячейки из столбца column.
        /// Для приведения значения в тип T используется контрукция: (T)Convert.ChangeType(CellsValues[index], typeof(T));
        /// </summary>
        /// <param name="column">столбец из коллекции Columns</param>
        /// <exception cref="ArgumentException">Если имя столбец отсутствует в коллекии Columns</exception>
        /// <returns></returns>
        public T GetCellValue<T>(TableColumn column)
        {
            for (int i = 0; i < Columns.Count; i++)
            {
                if (Columns[i].Equals(column))
                {
                    return GetCellValue<T>(i);
                }
            }
            throw new ArgumentException("Column not in Columns collection");
        }
        /// <summary>
        /// Получить значение ячейки из столбца с номером index в  Columns
        /// Для приведения значения в тип T используется контрукция: (T)Convert.ChangeType(CellsValues[index], typeof(T));
        /// </summary>
        /// <param name="columnName">Номер столбца из коллекции Columns</param>
        /// <returns></returns>
        public T GetCellValue<T>(int index)
        {
            return TypeConverter.From<T>(CellsValues[index]);
            //if (CellsValues[index] == null)
            //    return default(T);
            //Type unNullableType = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
            //return (T)Convert.ChangeType(CellsValues[index], unNullableType);

            //if (CellsValues[index] == null)
            //    return default(T);
            //Type unNullableType = typeof(T);  //Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);

            //var converter = TypeDescriptor.GetConverter(unNullableType);
            //return converter.CanConvertFrom(CellsValues[index].GetType()) ?
            //    (T)converter.ConvertFrom(CellsValues[index]) :
            //    throw new InvalidCastException("Can't format " + CellsValues[index].ToString() + "   To " + typeof(T).ToString());

        }

        public IEnumerator<object> GetEnumerator()
        {
            return CellsValues.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
