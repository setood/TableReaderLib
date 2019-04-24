using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReaderForTableReaderLib
{
    internal class MSExcelWorker
    {
        static MSExcelWorker singletonInstance;
        static Excel._Application App { get; set; }
        static Excel._Worksheet WorkSheet { get; set; }
        static Excel._Workbook WorkBook { get; set; }
        //static Dictionary<string, Excel._Workbook> WorkBoooks { get; set; }
        //static Dictionary<string, Excel._Worksheet> WorkSheets { get; set; }
        static string FilePath { get; set; }
        static string SheetName { get; set; }
        static bool IsInit = false;
        public bool IsOpened => IsInit;
        private MSExcelWorker() { }

        private static void Init()
        {
            if (IsInit == true)
                return;
            App = new Excel.Application();
            App.Visible = false;

            App.DisplayAlerts = false;
#if DEBUG
            App.Visible = true;
            App.DisplayAlerts = true;
#endif
            WorkBook = (Excel._Workbook)(App.Workbooks.Open(FilePath));
            if (SheetName == null)
                WorkSheet = WorkBook.Sheets.Cast<Excel._Worksheet>().FirstOrDefault();// первый лист
            else
                WorkSheet = WorkBook.Sheets.Cast<Excel._Worksheet>().FirstOrDefault(r => r.Name == SheetName);
            if (WorkSheet == null)
            {
                WorkSheet = (Excel._Worksheet)WorkBook.Sheets.Add();
                WorkSheet.Name = SheetName ?? "1";
            }
            IsInit = true;
        }

        /// <summary>
        /// Идексы от нуля
        /// </summary>
        /// <param name="startRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="endColumnIndex"></param>
        /// <returns></returns>
        public CellsRange GetCellsRange(int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            if (FilePath == null || SheetName == null)
                throw new InvalidOperationException("Before calls this method, you must call Open(filePath, sheetName) method.");
            Init();
            startRowIndex++;
            startColumnIndex++;
            endRowIndex++;
            endColumnIndex++;
            if (startRowIndex == endRowIndex && startColumnIndex == endColumnIndex)
            {
                ///Если вместо диапазона ячеекуказан адрес одной ячейки, возвращается не массив значений а  значение одной ячейки, что ломает логику.
                endColumnIndex++;
            }

            var startCell = WorkSheet.Cells[startRowIndex, startColumnIndex];
            var endCell = WorkSheet.Cells[endRowIndex, endColumnIndex];
           
            return new CellsRange( WorkSheet.Range[startCell, endCell].Value2);
        }

        public static MSExcelWorker GetInstance()
        {
            if (singletonInstance == null)
                singletonInstance = new MSExcelWorker();
            return singletonInstance;
        }
        public void Open(string filePath, string sheetName = null)
        {
            if (FilePath == filePath && SheetName == sheetName)
                return;

            if (IsInit)
                throw new InvalidOperationException("WorkSheet already opened. Call Close() method before.");

            FilePath = filePath;
            SheetName = sheetName;
            Init();
        }

        public void Close()
        {
            Dispose();
        }


        public void Dispose()
        {
            //GC.Collect();
            //GC.WaitForPendingFinalizers();

            if (WorkSheet != null)
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(WorkSheet);
                WorkSheet = null;
            }
            if (WorkBook != null)
            {
                WorkBook.Close();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(WorkBook);
                WorkBook = null;
            }
            if (App != null)
            {
                App.Visible = true;
                App.DisplayAlerts = true;
                App.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(App);
                App = null;
            }
            IsInit = false;
            SheetName = null;
            FilePath = null;

        }

        

    }
}
