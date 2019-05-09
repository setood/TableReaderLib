using ExcelReaderForTableReaderLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReaderLib;
using ReadersForTableReaderLib;

namespace ConsoleAppDotNet
{
    class Program
    {
        static void Main(string[] args)
        {
            // testExcelRangeClient();
            testCsvReader();
        }

        private static void testExcelRangeClient()
        {
            using (MSExcelSourceReader rdr = new MSExcelSourceReader(@"C:\Users\sevik\source\repos\TableReaderLib\DemoConsoleAppDotNet\TextSample.xlsx", "TextSample", 3, true))
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

        private static void testCsvReader()
        {
            var reader = new CsvSourceReader(@"D:\myCsv.csv", "|");
            var columns = new TableColumn[]
            {
                new TableColumn(){IndexInSource  = 0 },
                new TableColumn(){IndexInSource  = 1 },
                new TableColumn(){IndexInSource  = 2 },
                new TableColumn(){IndexInSource  = 3 },
            };

            var table = new TableReader(reader, columns);
            table.StartRow = 2; //пустые строки перед началом нашей таблицы в файле
            table.IsFirstRowHeaders = true; //таблица имеет строку заголовков





            string headers = string.Join(";\t",table.SourceHeaders);
            var row = table.ToArray()[2]; 
            string rowInfo =
                $"column 1 value:{row.GetCellValue<int>(0)}, type:{row.GetCellValue<int>(0).GetType()}" + Environment.NewLine +
                $"column 2 value:{row.GetCellValue<string>(1)}, type:{row.GetCellValue<string>(1).GetType()}" + Environment.NewLine +
                $"column 3 value:{row.GetCellValue<DateTime>(2)}, type:{row.GetCellValue<DateTime>(2).GetType()}" + Environment.NewLine +
                $"column 4 value:{row.GetCellValue<TimeSpan>(3)}, type:{row.GetCellValue<TimeSpan>(3).GetType()}" + Environment.NewLine;

            Console.WriteLine(headers);
            Console.WriteLine(rowInfo);
            Console.ReadLine();

        }

    }

}
