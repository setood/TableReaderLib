using System;
using System.Collections.Generic;
using System.Text;

namespace TableReaderLib
{
    /// <summary>
    /// Реаизация интерфейся позволяет использовать объекты чтения различных источников единым образом.
    /// Так как объекты реализующие ISourceReader могут создаваться совершенно по разному, а вызывающемо коду необходимо создавать их копии (для повторного чтения того же источника)
    ///     Использован создающий метод CreateReaderClone.
    /// IEnumerator<TableRow> - позовляет получить все строки в соответствии с параметрами метода CreateReaderClone.
    /// </summary>
    public interface ISourceReader : IEnumerator<TableRow>
    {
        /// <summary>
        /// Метод создает копию объекта чтения источника, и изменяет указанные в параметрах метода свойства копии.
        /// </summary>
        /// <param name="columns">Коллекция столбцов, которую необходимо считывать из источника</param>
        /// <param name="isFirstRowHeaders">Является ли первая строка Заголовками столбцов.</param>
        /// <param name="startRow">Указывает на начало таблицы. Например когда в первых строках источника идет описание таблицы, а ниже сама таблица</param>
        /// <param name="skippedRows">Сколько строк таблицы необходимо пропустить. Например если они были прочитаны ранее</param>
        /// <param name="takeRows">Сколько строк таблицы нужно прочитать. Null если необходимо считывать до конца таблицы.</param>
        /// <returns></returns>
        ISourceReader CreateReaderClone(IEnumerable<TableColumn> columns, bool isFirstRowHeaders, int startRow, int skippedRows, int? takeRows);
    }
}
