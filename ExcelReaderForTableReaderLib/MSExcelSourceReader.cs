using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TableReaderLib;

namespace ExcelReaderForTableReaderLib
{
    public class MSExcelSourceReader : ISourceReader
    {
        #region especiality class code

        //public string FilePath { get; private set; }
        //public string SheetName { get; private set; }
        private int CurrentRowIndex { get; set; }
        private int RowsInCacheWindow { get; set; }
        /// <summary>
        /// Первая строка которую необходимо считывать
        /// </summary>
        private int FirstRowInTableIndex => StartRow +
            (IsFirstRowHeaders == true ? 1 : 0) +
            SkippedRows;
        /// <summary>
        /// Последняя доступная для чтения строка. ниже неё нельзя помещать фокус
        /// </summary>
        private int LastTakenRowIndex => TakeRows == null ? int.MaxValue : FirstRowInTableIndex + TakeRows.Value;
        private int LastColumnIndex => (Columns == null || Columns.Count() == 0) ? 0 : Columns.Max(c => c.IndexInSource);

        private MSExcelWorker excelWorker { get; set; }
        private ReadingBooster Booster { get; set; }
        private bool IsEmptyRowIsTableEnd {get; set;}

        public MSExcelSourceReader(string filePath, string sheetName, int rowsInCacheWindow, bool isEmptyRowIsTableEnd)
        {
            //FilePath = filePath;
            //SheetName = sheetName;
            IsEmptyRowIsTableEnd = isEmptyRowIsTableEnd;
            RowsInCacheWindow = rowsInCacheWindow;
            excelWorker = new MSExcelWorker();
            excelWorker.Open( filePath, sheetName);
            Columns = new TableColumn[0];
            Reset();
        }

        private MSExcelSourceReader(MSExcelWorker excelWorker, int rowsInCacheWindow)
        {
            RowsInCacheWindow = rowsInCacheWindow;
            this.excelWorker = excelWorker;
            Columns = new TableColumn[0];
            Reset();
        }



        #endregion

        #region ISourceReader
        private int _startRow;
        private bool _isFirstRowHeaders;
        private int _skippedRows;
        private int? _takeRows;
        private IEnumerable<TableColumn> _columns;
        private int StartRow { get => _startRow; set { _startRow = value; Reset(); } }
        private bool IsFirstRowHeaders { get => _isFirstRowHeaders; set { _isFirstRowHeaders = value; Reset(); } }
        private int SkippedRows { get => _skippedRows; set { _skippedRows = value; Reset(); } }
        private int? TakeRows { get => _takeRows; set { _takeRows = value; Reset(); } }
        private IEnumerable<TableColumn> Columns { get => _columns; set { _columns = value; Reset(); } }
        public ISourceReader CreateReaderClone(IEnumerable<TableColumn> columns, bool isFirstRowHeaders, int startRow, int skippedRows, int? takeRows)
        {
            var thisCopy = new MSExcelSourceReader(this.excelWorker, this.RowsInCacheWindow);
            //thisCopy.FocusPosition = 0;
            //thisCopy.RowsInFocus = this.RowsInFocus;
            //thisCopy.CurrentRowInFocus = 0;
            //thisCopy.WorkRange = null;
            //thisCopy.excelWorker = this.excelWorker;
            thisCopy.IsEmptyRowIsTableEnd = this.IsEmptyRowIsTableEnd;
            thisCopy.Columns = columns;
            thisCopy.IsFirstRowHeaders = isFirstRowHeaders;
            thisCopy.StartRow = startRow;
            thisCopy.SkippedRows = skippedRows;
            thisCopy.TakeRows = takeRows;
            return thisCopy;
        }

        public TableRow Current
        {
            get
            {

                if (CurrentRowIndex < 0)
                {
                    return null;

                }

                var RowData = Booster.GetRow(CurrentRowIndex);
                TableRow row = new TableRow(this.Columns.ToArray(), RowData, ExcelTypeConverter.Converter);
                return row;

            }
        }

        object IEnumerator.Current => this.Current;



        //public bool MoveNext()
        //{
        //    ///если CurrentRowInFocus < RowsInFocus
        //    ///     сдвинуть и вернуть true
        //    ///иначе
        //    ///     можно свдвинуть фокус?
        //    ///        сдвигаем. перемещаем указатель. вернуть тру
        //    ///      иначе вернуть false
        //    //if (CurrentRowInFocus < RowsInFocus - 1 && WorkRange != null)
        //    //{
        //    //    CurrentRowInFocus++;
        //    //    tempCurrentRow = null;
        //    //    return true;
        //    //}
        //    //else
        //    //{
        //    //    bool isFocusMoved = MoveFocusNext();
        //    //    if (isFocusMoved)
        //    //    {
        //    //        CurrentRowInFocus++;
        //    //        tempCurrentRow = null;
        //    //        return true;
        //    //    }
        //    //    else
        //    //        return false;
        //    //}

        //}

        public void Reset()
        {
            int preInitIndex = -1;

            CurrentRowIndex = preInitIndex + FirstRowInTableIndex;
            Booster = new ReadingBooster(this.excelWorker.WorkSheet, this.LastColumnIndex);
            

        }

        /// <summary>
        /// Сдвигает фокус в следующее положение
        /// </summary>
        

        public bool MoveNext()
        {
            if (CurrentRowIndex + 1 >= LastTakenRowIndex)
                return false;

            if (IsEmptyRowIsTableEnd && CurrentRowIndex >= 0 && this.Current.All(c => c == null))
            {
                return false;
            }
            CurrentRowIndex++;
            return true;
         }
        #endregion
        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls
       

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    //excelWorker.Dispose();
                    excelWorker = null;
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                


                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~MSExcelSourceReader() {
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
