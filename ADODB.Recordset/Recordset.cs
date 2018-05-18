using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADODB
{
    public class Recordset
    {
        private ADODB.Connection innerADOConnection = null;

        public System.Data.IDbCommand innerCommand { get; private set; }

        public System.Data.IDataReader innerReader { get; private set; }

        public bool _EOF = true;
        public bool EOF
        {
            get
            {
                OpenReader();
                
                return _EOF;
            }
        }

        //For future alteration, this going to be the Type injector
        private void CreateCommand(string sqlCommand)
        {
            innerCommand = new System.Data.SqlClient.SqlCommand(sqlCommand, (SqlConnection)innerADOConnection.innerConnection);
        }
        //Update reader and flag
        private bool ReaderRead()
        {
            _EOF = !innerReader.Read();
            return !_EOF;
        }
        private void OpenReader()
        {
            if (innerReader != null && !innerReader.IsClosed)
                return;

            if (innerADOConnection == null && innerADOConnection.innerConnection == null)
                throw new NullReferenceException("Connection not Initialized.");

            innerADOConnection.innerConnection.Open();

            if (innerADOConnection.innerConnection.State == System.Data.ConnectionState.Closed || innerADOConnection.innerConnection.State == System.Data.ConnectionState.Broken)
                throw new InvalidOperationException("Connection is not in correte State.");

            if (innerCommand == null)
                throw new NullReferenceException("Command not Initialized.");

            innerReader = innerCommand.ExecuteReader();

            if (!ReaderRead())
                throw new InvalidOperationException("Reader not executed.");

        }

        public void Open(string sqlCommand, ADODB.Connection connection)
        {
            if (connection != null || innerADOConnection == null)
                innerADOConnection = connection;

            CreateCommand(sqlCommand);
        }

        public RecordsetItem fields(int index)
        {
            OpenReader();

            var item = new RecordsetItem(innerReader.GetName(index), innerReader[index]);

            return item;
        }
        public RecordsetItem fields(string index)
        {
            OpenReader();

            //Just to make sure 
            var item = new RecordsetItem(innerReader.GetName(innerReader.GetOrdinal(index)), innerReader[index]);

            return item;
        }
        public IEnumerable<RecordsetItem> fields()
        {
            OpenReader();

            for (int i = 0; i < innerReader.FieldCount; i++)
            {
                var item = new RecordsetItem(innerReader.GetName(i), innerReader[i]);

                yield return item;
            }

            yield break;
        }

        public void MoveNext()
        {
            ReaderRead();
            //_EOF = !innerReader.Read();
        }
    }
}
