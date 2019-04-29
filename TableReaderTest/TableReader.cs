using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TableReaderLib;
using System.Linq;
using System.Text;

namespace TableReaderTest
{
    [TestClass]
    public class TableTextReader
    {
        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext { get; set; }
        static Dictionary<string, TableReaderLib.TableReader> _tableTextReadersCollection;

        [ClassInitialize]
        public static void CreateReaders(TestContext testContext)
        {
            CleanAll();
            var readersCollection = new Dictionary<string, ISourceReader>()
            {
                { "ReadersForTableReaderLib.CsvSourceReader + .txt", new ReadersForTableReaderLib.CsvSourceReader(@"C:\Users\sevik\source\repos\TableReaderLib\TableReaderTest\ForSourceTest\TextSample.txt", "|") },
             { "ReadersForTableReaderLib.CsvSourceReader + .csv", new ReadersForTableReaderLib.CsvSourceReader(@"C:\Users\sevik\source\repos\TableReaderLib\TableReaderTest\ForSourceTest\TextSample.csv", ";") },
                {"ExcelReaderForTableReaderLib.MSExcelSourceReader + .xlsx",
                    new ExcelReaderForTableReaderLib.MSExcelSourceReader(
                        filePath: @"C:\Users\sevik\source\repos\TableReaderLib\TableReaderTest\ForSourceTest\TextSample.xlsx",
                        sheetName: "TextSample",
                        rowsInCacheWindow: 10,
                        isEmptyRowIsTableEnd: true) }
            };
            _tableTextReadersCollection = new Dictionary<string, TableReaderLib.TableReader>();
            foreach (var rdr in readersCollection)
            {
                    var tRdr =
                        new TableReaderLib.TableReader(
                            rdr.Value,
                            TestTableModel.Columns);
                    tRdr.IsFirstRowHeaders = true;
                tRdr.TakeRows = 100;
                    _tableTextReadersCollection.Add(rdr.Key, tRdr);
            }

        }


        [ClassCleanupAttribute]
        public static void CleanAll()
        {
            var excels = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            for (int i = 0; i< excels.Length; i++)
            {
                excels[i].Kill();
            }
        }



        
        public static void ResetReadersState()
        {
            var readers = _tableTextReadersCollection.Select(r => r.Value).ToArray();
            for (int i = 0; i < readers.Count(); i++)
            {
                var r = readers[i];
                r.IsFirstRowHeaders = true;
                r.TakeRows = 60;
                r.SkippedRows = 0;
                r.StartRow = 0;
            }
        }

        [TestMethod]
        public void RowsCount()
        {
            ResetReadersState();
            //System.Windows.MessageBox.Show("RowsCount start");
            var model = TestTableModel.Data.ToArray();
            foreach (var t in _tableTextReadersCollection)
            {
                //System.Windows.MessageBox.Show($"t is {t.Key}\r\n");
                int modelRowsCount = model.Count();
                int rdrRowsCount = t.Value.Count();
                Assert.AreEqual<int>(modelRowsCount, rdrRowsCount);
                //System.Windows.MessageBox.Show($"target:{modelRowsCount} \t\t result:{rdrRowsCount}");
            }
        }

        [TestMethod]
        public void TestCellsReading()
        {
            var model = TestTableModel.Data.ToArray();
            foreach (var t in _tableTextReadersCollection)
            {
                var cells = t.Value.Select(r => //new TestTableModel.Row()
                {
                    var row = new TestTableModel.Row();
                    row.Col0 = r.GetCellValue<string>(0);
                    //if (t.Key.ToLower().Contains("excel"))
                    //{
                    //    var val =  r.GetCellValue<double?>(1);
                    //    row.Col1 = (int)val;
                    //}
                    //else
                        row.Col1 = r.GetCellValue<int?>(1);
                    row.Col2 = r.GetCellValue<double?>(2);
                    row.Col3 = r.GetCellValue<DateTime?>(3);
                    row.Col4 = r.GetCellValue<TimeSpan?>(4);
                    return row;
                }).ToArray();
                for (int i = 0; i < model.Length; i++)
                {
                    Assert.AreEqual(model[i][0]??"", cells[i].Col0??"");
                    Assert.AreEqual(model[i][1], cells[i].Col1);
                    Assert.AreEqual(model[i][2], cells[i].Col2);
                    Assert.AreEqual(model[i][3], cells[i].Col3);
                    Assert.AreEqual(model[i][4], cells[i].Col4);
                }
            }
        }
        [TestMethod]
        public void TestSourceHeaders()
        {
            ResetReadersState();
            foreach (var t in _tableTextReadersCollection)
            {
                foreach (var col in TestTableModel.Columns)
                {
                    Assert.AreEqual(col.Name, t.Value.SourceHeaders[col.IndexInSource]);
                }
            }
        }
        [TestMethod]
        public void TestStartRow()
        {
            ResetReadersState();
            foreach (var t in _tableTextReadersCollection)
            {
                
                t.Value.StartRow = 1;
                ///первую строку скипнули, и в заголовки попали значения которые в модели данных являются стройкой.
                Assert.AreEqual(t.Value.SourceHeaders[0], TestTableModel.Data[0][0].ToString());
                Assert.AreEqual(t.Value.SourceHeaders[1], TestTableModel.Data[0][1].ToString());

                t.Value.StartRow = 2;
                Assert.AreEqual(t.Value.SourceHeaders[0], TestTableModel.Data[1][0].ToString());
            }


        }
        [TestMethod]
        public void TestIsFirstRowHeaders()
        {
            ResetReadersState();
            foreach (var t in _tableTextReadersCollection)
            {
                t.Value.IsFirstRowHeaders = true;
                Assert.AreEqual(t.Value.SourceHeaders[0], TestTableModel.Columns[0].Name);
                Assert.AreEqual(t.Value.SourceHeaders[1], TestTableModel.Columns[1].Name);

                t.Value.IsFirstRowHeaders = false;
                Assert.IsTrue(t.Value.SourceHeaders.Count() == 0);

            }
        }
        [TestMethod]
        public void TestSkippedRows()
        {
            //Заголовки должны быть правильными.
            //а первая строка должна быть с учетом Skip'a
            foreach (var t in _tableTextReadersCollection)
            {
                t.Value.StartRow = 0;
                t.Value.SkippedRows = 1;
                t.Value.IsFirstRowHeaders = true;
                Assert.AreEqual(t.Value.SourceHeaders[0], TestTableModel.Columns[0].Name);
                Assert.AreEqual(t.Value.SourceHeaders[1], TestTableModel.Columns[1].Name);

                var cells = t.Value.Select(r => //new TestTableModel.Row()
                {
                    var row = new TestTableModel.Row();
                    row.Col0 = r.GetCellValue<string>(0);
                    row.Col1 = r.GetCellValue<int?>(1);
                    row.Col2 = r.GetCellValue<double?>(2);
                    row.Col3 = r.GetCellValue<DateTime?>(3);
                    row.Col4 = r.GetCellValue<TimeSpan?>(4);
                    return row;
                }).ToArray();
                var model = TestTableModel.Data.ToArray();
                for (int i = 0; i < model.Length - t.Value.SkippedRows; i++)
                {
                    var v1 = model[i + t.Value.SkippedRows][0];
                    var v2 = cells[i].Col0;

                    var v3 = model[i + t.Value.SkippedRows][1];
                    var v4 = cells[i].Col1;

                    var v5 = model[i + t.Value.SkippedRows][2];
                    var v6 = cells[i].Col2;

                    Assert.AreEqual(model[i+ t.Value.SkippedRows][0] ?? "", cells[i].Col0 ?? "");
                    Assert.AreEqual(model[i+ t.Value.SkippedRows][1], cells[i].Col1);
                    Assert.AreEqual(model[i+ t.Value.SkippedRows][2], cells[i].Col2);
                    Assert.AreEqual(model[i+ t.Value.SkippedRows][3], cells[i].Col3);
                    Assert.AreEqual(model[i+ t.Value.SkippedRows][4], cells[i].Col4);
                }

            }
        }
        [TestMethod]
        public void TestTakeRows()
        {
            ResetReadersState();
            foreach (var t in _tableTextReadersCollection)
            {
                ///проверить количество строк.
                ///проверить содержимое
                ///проверить если TakeRows больше чем строк в источнике
                t.Value.TakeRows = 1;
                t.Value.StartRow = 0;
                t.Value.SkippedRows = 0;
                var rows = t.Value.ToArray();
                var v01 = t.Value.Count();
                var v02 = t.Value.TakeRows;
                Assert.AreEqual(t.Value.Count(), t.Value.TakeRows);
                t.Value.TakeRows = TestTableModel.Data.Count() + 10;
                Assert.AreEqual(t.Value.Count(), TestTableModel.Data.Count());
            }
        }
    }
}
