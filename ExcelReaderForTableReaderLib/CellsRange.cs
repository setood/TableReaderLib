using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderForTableReaderLib
{
    /// <summary>
    /// Copy range of values from Excel gives non-Zero based array like  ExcelRange[1..4,1..10].
    /// This Class is wrapper that makes range of values zero based.
    /// </summary>
    class CellsRange
    {
        object[,] ExcelRange { get; set; }
        public object this[int row, int column]
        {
            get => ExcelRange[row + 1, column + 1];
        }



        public CellsRange(object[,] excelRange)
        {
            this.ExcelRange = excelRange;
        }
        public CellsRange(_Worksheet worksheet, int startRow, int rowsCount, int startColumnIndex, int columnsCount)
        {
            var startCell = worksheet.Cells[startRow + 1, startColumnIndex + 1];
            var endCell = worksheet.Cells[startRow + 1 + rowsCount, startColumnIndex + 1 + columnsCount];
            this.ExcelRange = worksheet.Range[startCell, endCell].Value2;

        }

        public object[] GetRow(int rowIndex)
        {
            return GetMultiDementialArrayRow(ExcelRange, rowIndex);
        }
        
        private T[] GetMultiDementialArrayRow<T>(T[,] matrix, int rowIndex)
        {
            return Enumerable.Range(0, matrix.GetLength(1))
                    .Select(x => matrix[rowIndex, x])
                    .ToArray();
        }
    }
}
