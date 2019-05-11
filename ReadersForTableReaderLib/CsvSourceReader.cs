using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Linq;
using TableReaderLib;

namespace ReadersForTableReaderLib
{
    public class CsvSourceReader : ISourceReader
    {
        string filePath;
        Encoding encoding;
        protected StreamReader sr;
        protected string[] currentData;
        protected string[] _splitters;
        int readedRows;
        IEnumerable<TableColumn> _columns;
        public IEnumerable<TableColumn> Columns
        {
            get => _columns;
            set => _columns = value;

        }
        bool _isFirstRowHeaders;
        bool IsFirstRowHeaders
        {
            get => _isFirstRowHeaders;
            set
            {
                if (value == _isFirstRowHeaders)
                    return;
                _isFirstRowHeaders = value;
                Reset();
            }
        }
        int _startRow;
        int StartRow
        {
            get => _startRow;
            set
            {
                if (_startRow != value)
                {
                    _startRow = value;
                    Reset();
                }
            }
        }
        int _skippedRows;
        int SkippedRows
        {
            get => _skippedRows;
            set
            {
                if (_skippedRows != value)
                {
                    _skippedRows = value;
                    Reset();
                }
            }
        }
        int? _takeRows;
        int? TakeRows
        {
            get => _takeRows;
            set
            {
                if (_takeRows != value)
                {
                    _takeRows = value;
                    Reset();
                }
            }
        }
        
        public CsvSourceReader(string filePath, string splitter = null, Encoding encoding = null)
        {
            this.filePath = filePath;
            this.encoding = encoding;
            sr = new StreamReader(filePath, encoding ?? Encoding.Default);
            if (splitter == null)
                splitter = "|";
            _splitters = new string[] { splitter };
            
        }

         bool Read()
        {
            if (TakeRows != null && TakeRows <= readedRows)
                return false;
            string temp = sr.ReadLine();
            if (temp == null)
                return false;
            currentData = temp.Split(_splitters, StringSplitOptions.None).Select(s=>string.IsNullOrEmpty(s)?null : s).ToArray();
            readedRows++;
            return true;
        }

        public  T GetColumnValue<T>(int columnIndex, bool returnDefaultTypeValueIfException = false)
        {
#if DEBUG
            try
            {
#endif
                IFormatProvider frmt = CultureInfo.InvariantCulture;
                if (typeof(T) == typeof(DateTime) || typeof(T) == typeof(DateTime?))
                    frmt = CultureInfo.CurrentCulture;
                return (T)Convert.ChangeType(currentData[columnIndex].Trim(), typeof(T), frmt);
#if DEBUG
            }
            catch (Exception ex)
            {
                throw;
            }
#endif
        }

        public ISourceReader CreateReaderClone(IEnumerable<TableColumn> columns, bool isFirstRowHeaders, 
            int startRow, int skippedRows, int? takeRows)
        {
            var rdr = new CsvSourceReader(this.filePath, this._splitters[0], this.encoding);
            rdr.Columns = columns;
            rdr.IsFirstRowHeaders = isFirstRowHeaders;
            rdr.StartRow = startRow;
            rdr.SkippedRows = skippedRows;
            rdr.TakeRows = takeRows;
            return rdr;
        }

        private void SkipStartRows()
        {
            for (int i = 0; i < this.StartRow; i++)
            {
                sr.ReadLine();
            }
        }

        private void SkipAndReadHeaders()
        {
            if (this.IsFirstRowHeaders == false)
                return;
            //this.Read();
            sr.ReadLine();
        }
  
        private void SkipSkippedRows()
        {
            for (int i = 0; i < this.SkippedRows; i++)
            {
                //this.Read();
                sr.ReadLine();
            }
        }



        public TableRow Current => new TableRow(this.Columns.ToArray(), this.Columns.Select(c => (object)this.currentData[c.IndexInSource]).ToArray());

        object IEnumerator.Current => this.Current;

        public bool MoveNext()
        {
            return this.Read();
        }

        public void Reset()
        {
            sr.DiscardBufferedData();
            sr.BaseStream.Position = 0;
            readedRows = 0;
            SkipStartRows();
            SkipAndReadHeaders();
            SkipSkippedRows();
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    
                    currentData = null;
                    sr.Close();
                    GC.Collect();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~SourceReader() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }

        
        #endregion
    }
}
