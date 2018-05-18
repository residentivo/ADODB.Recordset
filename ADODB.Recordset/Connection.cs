using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADODB
{
    public class Connection
    {
        //For future update
        public System.Data.IDbConnection innerConnection { get; private set; }

        //public string Provider { get; set; }

        public void Open(string stringConnection)
        {
            innerConnection = new System.Data.SqlClient.SqlConnection(stringConnection);        
        }
    }
}
