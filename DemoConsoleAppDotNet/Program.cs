using ExcelReaderForTableReaderLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReaderLib;

namespace ConsoleAppDotNet
{
    class Program
    {
        static void Main(string[] args)
        {
            testExcelRangeClient();

        }

        private static void testExcelRangeClient()
        {
            using (MSExcelSourceReader rdr = new MSExcelSourceReader(@"C:\Users\sevik\source\repos\TableReaderLib\DemoConsoleAppDotNet\TextSample.xlsx", "TextSample", 3))
            {
                TableColumn[] columns = new TableColumn[]
                {
                    new TableColumn(){Name = "1",  IndexInSource = 0},
                    new TableColumn(){Name = "2",  IndexInSource = 1},
                    new TableColumn(){Name = "3",  IndexInSource = 2},
                };


                var tr = new TableReader(rdr, columns);
                tr.IsFirstRowHeaders = true;
                tr.TakeRows = 10;
                var readQuery = tr.Select(r => new { r0 = r.GetCellValue<int?>(0), r1 = r.GetCellValue<double>(1), r2 = r.GetCellValue<string>(2) });
                var readedDat = readQuery.ToList<object>();
                //foreach (var r in tr)
                //{
                //    var rowValues = new { r0 = r.GetCellValue<int?>(0), r1 = r.GetCellValue<double>(1), r2 = r.GetCellValue<string>(2) };
                //    readedData.Add(rowValues);
                //}
            }
        }
    }
}
