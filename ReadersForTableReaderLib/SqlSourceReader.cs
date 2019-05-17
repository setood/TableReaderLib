using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using TableReaderLib;
using System.Globalization;

namespace ReadersForTableReaderLib
{
    public class SqlSourceReader : LineReader
    {

        private readonly string connectionString;
        private SqlConnection sqlConnection;
        private SqlDataReader dataReader;
        private string sqlQueryString;

        public SqlSourceReader(string connectionString, string sqlQueryString)
        {
            this.connectionString = connectionString;
            this.sqlQueryString = sqlQueryString;

            sqlConnection = new SqlConnection(connectionString);

            SqlCommand command =
                new SqlCommand(sqlQueryString, sqlConnection);
            //sqlConnection.Open();
            //dataReader = command.ExecuteReader();
            //Reset();
        }

        protected override bool ReadNextLine()
        {
            var result = dataReader.Read();
            if (result == false)
                return false;
            List<object> values = new List<object>();
            foreach (var c in Columns)
            {
                var val = dataReader.GetValue(c.IndexInSource);
                if (val == DBNull.Value)
                    val = null;
                values.Add(val);
            }
            base.currentData = values.ToArray();
            return true;
        }

        protected override void ResetSource()
        {
            dataReader?.Close();
            SqlCommand command =
                new SqlCommand(sqlQueryString, sqlConnection);
            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();
            dataReader = command.ExecuteReader();
            //if (IsFirstRowHeaders)
            //    dataReader.Read();
        }


        public override T GetColumnValue<T>(int columnIndex, bool returnDefaultTypeValueIfException = false)
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

        public override ISourceReader CreateReaderClone(IEnumerable<TableColumn> columns, bool isFirstRowHeaders, int startRow, int skippedRows, int? takeRows)
        {
            var sqlSourceReaderClone = new SqlSourceReader(this.connectionString, this.sqlQueryString);
            sqlSourceReaderClone.Columns = columns;
            sqlSourceReaderClone.IsFirstRowHeaders = isFirstRowHeaders;

            sqlSourceReaderClone.StartRow = startRow;
            sqlSourceReaderClone.SkippedRows = skippedRows;
            sqlSourceReaderClone.TakeRows = takeRows;
            sqlSourceReaderClone.Reset();
            //if (isFirstRowHeaders)
            //    sqlSourceReaderClone.ReadNextLine();
            return sqlSourceReaderClone;
        }
        
        protected override void OnDispose()
        {
            dataReader.Close();
            sqlConnection.Dispose();
            dataReader = null;
            sqlConnection = null;
        }
    }
}
