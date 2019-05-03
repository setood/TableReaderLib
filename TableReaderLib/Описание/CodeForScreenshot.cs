using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableReaderLib.Описание
{
    /// <summary>
    /// код нужен для создания скриншотов и не несет никакой смысловой нагрузки
    /// </summary>
    class CodeForScreenshot
    {
        dynamic data;
        interface _WorkSheet { dynamic[,] Cells { get; set; } }

        internal class MSExcelSourceReader : ISourceReader
        {
            private string filePath;

            public MSExcelSourceReader(string filePath)
            {
                this.filePath = filePath;
            }

            public TableRow Current => throw new NotImplementedException();

            object IEnumerator.Current => throw new NotImplementedException();

            public ISourceReader CreateReaderClone(IEnumerable<TableColumn> columns, bool isFirstRowHeaders, int startRow, int skippedRows, int? takeRows)
            {
                throw new NotImplementedException();
            }

            public void Dispose()
            {
                throw new NotImplementedException();
            }

            public bool MoveNext()
            {
                throw new NotImplementedException();
            }

            public void Reset()
            {
                throw new NotImplementedException();
            }
        }
        
        void ReadThisTable(_WorkSheet workSheet)
        {
            for (int rowIndex = 0; rowIndex < 30; rowIndex++)
            {
                if (workSheet.Cells[rowIndex, 0].Value == "Нужно")
                {
                    for (int colIndex = 0; colIndex < 5; colIndex++)
                    {
                        data[rowIndex][colIndex] = workSheet.Cells[rowIndex, colIndex].Value;
                    }
                }
                
            }
        }


        void ReadThisTableSimple(string filePath)
        {
            var excelReader = new MSExcelSourceReader(filePath);
            var isNeedColumn = new TableColumn() { Name = "Нужна ли строка", IndexInSource = 0 };
            var dataColumn = new TableColumn() { Name = "данные", IndexInSource = 1 };
            var columns = new TableColumn[] {isNeedColumn, dataColumn};
            var tableReader = new TableReader(excelReader, columns);

            data = tableReader.Where(row=>row.GetCellValue<string>(0) == "Нужно");
        }


    }
}
