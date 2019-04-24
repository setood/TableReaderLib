using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using TableReaderLib;

namespace ReadersForTableReaderLib
{
    public class SqlSourceReader : ISourceReader
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
            Reset();
        }

        #region ISourceReader implementation
        IEnumerable<TableColumn> columns;
        bool isFirstRowHeaders;
        int startRow;
        int skippedRows;
        int? takeRows;
        
            
        public TableRow Current
        {
            get
            {
                List<object> values = new List<object>();
                foreach (var c in columns)
                {
                    var val = dataReader.GetValue(c.IndexInSource);
                    if (val == DBNull.Value)
                        val = null;
                    values.Add(val);
                }
                var newRow = new TableRow(columns.ToArray(), values);
                return newRow;
            }

        }

        object IEnumerator.Current => this.Current;

        public ISourceReader CreateReaderClone(IEnumerable<TableColumn> columns, bool isFirstRowHeaders, int startRow, int skippedRows, int? takeRows)
        {
            var sqlSourceReaderClone = new SqlSourceReader(this.connectionString, this.sqlQueryString);
            sqlSourceReaderClone.columns = columns;
            sqlSourceReaderClone.isFirstRowHeaders = isFirstRowHeaders;
            sqlSourceReaderClone.startRow = startRow;
            sqlSourceReaderClone.skippedRows = skippedRows;
            sqlSourceReaderClone.takeRows = takeRows;
            return sqlSourceReaderClone;
        }

        public bool MoveNext()
        {
            return dataReader.Read();
        }

        public void Reset()
        {
            dataReader?.Close();
            SqlCommand command =
                new SqlCommand(sqlQueryString, sqlConnection);
            if (sqlConnection.State != ConnectionState.Open)
                sqlConnection.Open();
            dataReader = command.ExecuteReader();
        }
        #endregion
        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls
        

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                dataReader.Close();
                sqlConnection.Dispose();
                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~SqlSourceReader() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}
