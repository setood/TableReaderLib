using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Linq;
using TableReaderLib;

namespace ReadersForTableReaderLib
{
    public class CsvSourceReader : LineReader
    {
        string filePath;
        Encoding encoding;
        protected StreamReader sr;
        protected string[] _splitters;
        
        
        public CsvSourceReader(string filePath, string splitter = null, Encoding encoding = null)
        {
            this.filePath = filePath;
            this.encoding = encoding;
            sr = new StreamReader(filePath, encoding ?? Encoding.Default);
            if (splitter == null)
                splitter = "|";
            _splitters = new string[] { splitter };
            //needToReset = true;
        }

        protected override bool ReadNextLine()
        {
            string temp = sr.ReadLine();
            if (temp == null)
                return false;
            currentData = temp.Split(_splitters, StringSplitOptions.None).Select(s => string.IsNullOrEmpty(s) ? null : s).ToArray();
            return true;
        }
        public override  T GetColumnValue<T>(int columnIndex, bool returnDefaultTypeValueIfException = false)
        {
#if DEBUG
            try
            {
#endif
                IFormatProvider frmt = CultureInfo.InvariantCulture;
                if (typeof(T) == typeof(DateTime) || typeof(T) == typeof(DateTime?))
                    frmt = CultureInfo.CurrentCulture;
                return (T)Convert.ChangeType(((string)currentData[columnIndex]).Trim(), typeof(T), frmt);
#if DEBUG
            }
            catch (Exception ex)
            {
                throw;
            }
#endif
        }

        public override ISourceReader CreateReaderClone(IEnumerable<TableColumn> columns, bool isFirstRowHeaders, 
            int startRow, int skippedRows, int? takeRows)
        {
            var rdr = new CsvSourceReader(this.filePath, this._splitters[0], this.encoding);
            rdr.Columns = columns;
            rdr.IsFirstRowHeaders = isFirstRowHeaders;
            rdr.StartRow = startRow;
            rdr.SkippedRows = skippedRows;
            rdr.TakeRows = takeRows;
            return rdr;
        }

        
        protected override void ResetSource()
        {
            sr.DiscardBufferedData();
            sr.BaseStream.Position = 0;
        }

        protected override void OnDispose()
        {
            sr.Close();
            sr = null;
        }
       
    }
}
