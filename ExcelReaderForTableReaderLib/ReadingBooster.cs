using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReaderLib;

namespace ExcelReaderForTableReaderLib
{
    internal class ReadingBooster : IDisposable
    {
        #region public
        public ReadingBooster(_Worksheet worksheet, int lastColumnIndex, int rowsInWindows = 1000)
        {
            this.Sheet = worksheet;
            this.LastColumnIndex = lastColumnIndex;
            this.RowsInWindows = rowsInWindows;
        }

        public object[] GetRow(int rowIndex)
        {
            if (IsRowInCache(rowIndex) == false)
            {
                bool isMoveSuccess = MoveWindow(rowIndex);
                if (isMoveSuccess == false)
                    return null;
            }
            var rowInCache = GetRowIndexInCache(rowIndex);
            return Cache.GetRow(rowInCache); 
        }

       
        #endregion


        private _Worksheet Sheet { get; set; }
        private int LastColumnIndex { get; }
        private int RowsInWindows { get; }
        private CellsRange Cache { get; set; }
        private int CacheRowIndex { get; set; }
        private readonly int ExcelEndRow = 1048576;

        private bool IsRowInCache(int rowIndex) => 
            rowIndex >= CacheRowIndex && 
            rowIndex <= CacheRowIndex + RowsInWindows &&
            Cache !=null;
        private bool MoveWindow(int rowIndex)
        {
            //var newRowIndex = rowIndex + RowsInWindows;
            var newWindowsEndIndex = rowIndex + RowsInWindows;
            if (newWindowsEndIndex >= ExcelEndRow)
                return false;
            Cache = new CellsRange(Sheet, rowIndex, RowsInWindows, 0, LastColumnIndex);
            CacheRowIndex = rowIndex;
            return true;
        }
        private bool MoveWindowNext() => MoveWindow(RowsInWindows + CacheRowIndex);
        private int GetRowIndexInCache(int rowIndex) => rowIndex - CacheRowIndex;

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.
                Sheet = null;
                Cache = null;


                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~ReadingBooster() {
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
