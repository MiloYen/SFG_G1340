using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace G1340_SFG
{
    public class ConnectDB
    {
        SqlConnection cnn = null;
        public bool SQLIsConnect { get; set; }

        public void connect(string connectionString)
        {
            cnn = new SqlConnection(connectionString);
            try
            {
                cnn.Open();
                SQLIsConnect = true;
            }
            catch
            {
                SQLIsConnect = false;
            }
        }

        public void insert(string sqlComand)
        {
            SqlCommand cmd = new SqlCommand(sqlComand, cnn);
            SqlDataAdapter adapter = new SqlDataAdapter();
            adapter.InsertCommand = cmd;
            adapter.InsertCommand.ExecuteNonQuery();
            cmd.Dispose();
        }
    }
}
