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

        public string FilePath { get; private set; }
        public string SheetName { get; private set; }
        private int FocusPosition { get; set; }
        private int RowsInFocus { get; set; }
        private int CurrentRowInFocus { get; set; }
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
        /// <summary>
        /// на сколько строк можно перенести фокус
        /// </summary>
        private int FocusStepAvailable => Math.Min(LastTakenRowIndex - (FocusPosition + RowsInFocus), RowsInFocus);
        private CellsRange WorkRange { get; set; }
        private MSExcelWorker excelWorker { get; set; }
        public MSExcelSourceReader(string filePath, string sheetName, int rowsInFocus)
        {
            FilePath = filePath;
            SheetName = sheetName;
            RowsInFocus = rowsInFocus;
            excelWorker = MSExcelWorker.GetInstance();
            excelWorker.Open( filePath, sheetName);
            Columns = new TableColumn[0];
            Reset();
        }

        #endregion

        #region ISourceReader
        TableRow tempCurrentRow;
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
            var thisCopy = new MSExcelSourceReader(this.FilePath, this.SheetName, this.RowsInFocus);
            //thisCopy.FocusPosition = 0;
            //thisCopy.RowsInFocus = this.RowsInFocus;
            //thisCopy.CurrentRowInFocus = 0;
            //thisCopy.WorkRange = null;
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
                if (CurrentRowInFocus < 0)
                    throw new InvalidOperationException("call the MoveNext() method required");
                if (tempCurrentRow == null)
                {
                    List<object> values = new List<object>();
                    foreach (var c in this.Columns)
                    {
                        values.Add(WorkRange[CurrentRowInFocus, c.IndexInSource]);
                    }

                    tempCurrentRow = new TableRow(this.Columns.ToArray(), values);
                }
                return tempCurrentRow;
            }
        }

        object IEnumerator.Current => this.Current;



        public bool MoveNext()
        {
            ///если CurrentRowInFocus < RowsInFocus
            ///     сдвинуть и вернуть true
            ///иначе
            ///     можно свдвинуть фокус?
            ///        сдвигаем. перемещаем указатель. вернуть тру
            ///      иначе вернуть false
            if (CurrentRowInFocus < RowsInFocus - 1 && WorkRange != null)
            {
                CurrentRowInFocus++;
                tempCurrentRow = null;
                return true;
            }
            else
            {
                bool isFocusMoved = MoveFocusNext();
                if (isFocusMoved)
                {
                    CurrentRowInFocus++;
                    tempCurrentRow = null;
                    return true;
                }
                else
                    return false;
            }

        }

        public void Reset()
        {
            tempCurrentRow = null;
            FocusPosition = FirstRowInTableIndex - RowsInFocus;
            CurrentRowInFocus = RowsInFocus - 1; //при вызове MoveNext следующей строкой будет с нулевым индексом.

            WorkRange = null;

        }

        /// <summary>
        /// Сдвигает фокус в следующее положение
        /// </summary>
        private bool MoveFocusNext()
        {
            if (FocusStepAvailable <= 0)
                return false;
            CurrentRowInFocus = CurrentRowInFocus - FocusStepAvailable;
            FocusPosition = FocusPosition + FocusStepAvailable;
            int lastColumnInSource = Columns.Count() > 0 ?
                    Columns.Max(r => r.IndexInSource) :
                    0;
            WorkRange = excelWorker.GetCellsRange(FocusPosition, 0, FocusPosition + RowsInFocus, lastColumnInSource);
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
                    excelWorker.Dispose();
                    excelWorker = null;
                    WorkRange = null;
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
