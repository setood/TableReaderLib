using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReaderLib;

namespace ReadersForTableReaderLib
{
    public abstract class LineReader : ISourceReader
    {
        IEnumerable<TableColumn> _columns;

        bool needToReset = true;
        protected object[] currentData;
        int readedRows;
        bool _isFirstRowHeaders;
        protected bool IsFirstRowHeaders
        {
            get => _isFirstRowHeaders;
            set
            {
                if (value == _isFirstRowHeaders)
                    return;
                _isFirstRowHeaders = value;
                needToReset = true;
            }
        }
        int _startRow;
        protected int StartRow
        {
            get => _startRow;
            set
            {
                if (_startRow != value)
                {
                    _startRow = value;
                    needToReset = true;
                }
            }
        }
        int _skippedRows;
        protected int SkippedRows
        {
            get => _skippedRows;
            set
            {
                if (_skippedRows != value)
                {
                    _skippedRows = value;
                    needToReset = true;
                }
            }
        }
        int? _takeRows;
        protected int? TakeRows
        {
            get => _takeRows;
            set
            {
                if (_takeRows != value)
                {
                    _takeRows = value;
                    needToReset = true;
                }
            }
        }
    
        public IEnumerable<TableColumn> Columns
        {
            get => _columns;
            set => _columns = value;

        }
      
        protected abstract  bool ReadNextLine();
        public bool SkipLines(int count)
        {
            bool result = true;
            for (int i = 0; i < count && result == true; i++)
            {
                result = ReadNextLine();
            }
            return result;
        }


        public bool MoveNext()
        {
            if (needToReset == true)
                Reset();
            if (TakeRows != null && TakeRows + SkippedRows + StartRow + (IsFirstRowHeaders? 1:0) <= readedRows)
                return false;
            var result = ReadNextLine();
            if (result == true)
            {
                readedRows++;
                return true;
            }
            else return false;
        }

        public abstract T GetColumnValue<T>(int columnIndex, bool returnDefaultTypeValueIfException = false);

        public abstract ISourceReader CreateReaderClone(IEnumerable<TableColumn> columns, bool isFirstRowHeaders,
    int startRow, int skippedRows, int? takeRows);


        private void SkipStartRows()=>SkipLines(StartRow);
        

        private void SkipAndReadHeaders()
        {
            if (this.IsFirstRowHeaders == false)
                return;
            MoveNext();
        }

        private void SkipSkippedRows() => SkipLines(SkippedRows);

        public TableRow Current => new TableRow(this.Columns.ToArray(), this.Columns.Select(c => (object)this.currentData[c.IndexInSource]).ToArray());

        object IEnumerator.Current => this.Current;


        protected abstract void ResetSource();
        public void Reset()
        {
            needToReset = false;
            ResetSource();
            readedRows = 0;
            SkipStartRows();
            SkipAndReadHeaders();
            SkipSkippedRows();
            
        }

        #region IDisposable Support
        protected abstract void OnDispose();

        private bool disposedValue = false;
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    OnDispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.
                currentData = null;
                Columns = null;
                disposedValue = true;
            }
        }
        
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
        }


        #endregion
    }
}
