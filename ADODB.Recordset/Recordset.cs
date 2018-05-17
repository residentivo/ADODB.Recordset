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

        public System.Data.IDataReader innerReader { get; private set; }

        public void Open(string sqlCommand)
        {
            //throw new NotImplementedException();
        }

        public void Open(string sqlCommand, ADODB.Connection connection)
        {
            innerADOConnection = connection;
            //throw new NotImplementedException();
            CreateReader(sqlCommand);
        }

        private void CreateReader(string sqlCommand)
        {
            var cmd = new System.Data.SqlClient.SqlCommand(sqlCommand, (SqlConnection)innerADOConnection.innerConnection);

            if (cmd.Connection.State == System.Data.ConnectionState.Closed || cmd.Connection.State == System.Data.ConnectionState.Broken)
                cmd.Connection.Open();

            innerReader = cmd.ExecuteReader();
        }
    }
}
