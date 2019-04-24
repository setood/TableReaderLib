using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace TableReaderLib
{
    /// <summary>
    /// Объекты класса этого скрывают нюансы используемого источника данных, и создают впечетление работы с таблицей данных в оперативной памяти.
    /// Доступ к столбцам таблицы осуществляется через свойство Columns
    /// Доступ к строкам через интерфейс IEnumerable<TableRow>
    /// Доступ к ячейкам через строку из (IEnumerable<TableRow>) и столбец из Columns, 
    ///     Имя столбца (Имя в столбца в коллекции Columns. Имя столбца в источнике не учитывается) или 
    ///     номер столбца в Columns (Номер столбца в источнике, так же не учитывается)
    /// При изменениие каких нибудь свойств (например startRow), это затронет только новые вызовы GetEnumerator. Вызовы совершенные ранее это изменение не затронет.
    /// </summary>
    public sealed  class TableReader : IEnumerable<TableRow>
    {
        IReadOnlyList<TableColumn> _columns;
        /// <summary>
        /// Набор столбцов считываемых из источника
        /// </summary>
        public IReadOnlyList<TableColumn> Columns { get => _columns; set
            {
                if (_columns != value)
                {
                    _columns = value;
                    _sourceHeaders = null;
                }
            }
        }

        bool _isFirstRowHeaders;
        /// <summary>
        /// True означает, что первая  строка после StartRow является заголовками таблицы
        /// </summary>
        public bool IsFirstRowHeaders { get => _isFirstRowHeaders; set
            {
                if (_isFirstRowHeaders != value)
                {
                    _isFirstRowHeaders = value;
                    _sourceHeaders = null;
                }
            }
        }

        private IReadOnlyList<string> _sourceHeaders;
        /// <summary>
        /// Считывает заголовки из источника,  если IsFirstRowHeaders = True. иначе возвращает null
        /// </summary>
        public IReadOnlyList<string> SourceHeaders
        {
            get
            {
                if (IsFirstRowHeaders == false)
                    return new string[]{ };
                if (_sourceHeaders == null)
                {
                    var rdr = SourceReader.CreateReaderClone(Columns, false /*IsFirstRowHeaders*/, StartRow, 0, 1);
                    rdr.MoveNext();
                    string[] headers = new string[Columns.Count];
                    for (int i = 0; i < Columns.Count; i++)
                    {
                        headers[i] = rdr.Current.GetCellValue<string>(Columns[i].IndexInSource);
                    }
                    _sourceHeaders = headers;
                }
                return _sourceHeaders;
            }
        }

        int _startRow;
        /// <summary>
        /// Номер строки от которой начинается таблица. например, часто в первых строках идет описание таблицы, затем заголовки и сама таблица.
        /// </summary>
        public int StartRow { get => _startRow; set
            {
                if (_startRow != value)
                {
                    _startRow = value;
                    _sourceHeaders = null;
                }
            }
        }
        /// <summary>
        /// Сколько строк в таблице пропустить. Чтение строк начинается от StartRow(смешение таблицы в источнике) + 1(Если IsFirstRowHeaders = true) + SkippedRows
        /// </summary>
        public int SkippedRows { get; set; }
        /// <summary>
        /// Сколько строк нужно прочитать. Если читать до конца таблице TakeRows = null
        /// </summary>
        public int? TakeRows { get; set; }


        ISourceReader SourceReader { get; set; }

        public TableReader(ISourceReader sourceReader, IReadOnlyList<TableColumn> columns)
        {
            if (sourceReader == null || columns == null || columns.Count ==0 )
            {
                throw new ArgumentException("One of parameters is wrong. They are must be not null and columns not empty");
            }
            this.SourceReader = sourceReader;
            this.Columns = columns;
        }
        public IEnumerator<TableRow> GetEnumerator()
        {

            return SourceReader.CreateReaderClone(Columns, IsFirstRowHeaders, StartRow, SkippedRows, TakeRows);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
    
   
}
