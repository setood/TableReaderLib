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
    }
}
