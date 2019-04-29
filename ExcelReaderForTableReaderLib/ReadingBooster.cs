using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReaderLib;

namespace ExcelReaderForTableReaderLib
{
    internal class ReadingBooster
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

        //public T GetCellValue<T>(int columnIndex, int rowIndex)
        //{
        //    return GetRow(rowIndex).GetCellValue<T>(columnIndex);

        //}
        #endregion


        private _Worksheet Sheet { get; set; }
        private int LastColumnIndex { get; }
        private int RowsInWindows { get; }
        private CellsRange Cache { get; set; }
        private int CacheRowIndex { get; set; }
        private readonly int ExcelEndRow = 1048576;

        private bool IsRowInCache(int rowIndex) => rowIndex >= CacheRowIndex && rowIndex <= CacheRowIndex + RowsInWindows;
        private bool MoveWindow(int rowIndex)
        {
            //var newRowIndex = rowIndex + RowsInWindows;
            var newWindowsEndIndex = rowIndex + RowsInWindows;
            if (newWindowsEndIndex >= ExcelEndRow)
                return false;
            Cache = new CellsRange(Sheet, rowIndex, RowsInWindows, 1, LastColumnIndex);
            CacheRowIndex = rowIndex;
            return true;
        }
        private bool MoveWindowNext() => MoveWindow(RowsInWindows + CacheRowIndex);
        private int GetRowIndexInCache(int rowIndex) => rowIndex - CacheRowIndex;
    }
}
