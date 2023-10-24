using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Bai4_2.Data
{
    internal class DataBaseProcesser
    {
        string sqlConnectStr = "Data Source =(local);Initial Catalog = LTTrucQuan; User ID=sa; Password=abc123; Integrated Security = True";
        SqlConnection sqlConn = null;

        //PT mở kết nối
        void OpenConncect()
        {
            sqlConn = new SqlConnection(sqlConnectStr);
            if (sqlConn.State != ConnectionState.Open)
                sqlConn.Open();
        }
        //PT đóng kết nối
        void CloseConnect()
        {
            if (sqlConn.State != ConnectionState.Closed)
            {
                sqlConn.Close();
                sqlConn.Dispose();
            }
        }
        //PT thực hiện lệnh dạng insert, update, delete
        public void ChangeData(string sql)
        {
            OpenConncect();
            SqlCommand sqlCmm=new SqlCommand();
            sqlCmm.CommandText=sql;
            sqlCmm.Connection=sqlConn;
            sqlCmm.ExecuteNonQuery();
            CloseConnect();     
        }
        //PT thực hiện lệnh select
        public DataTable ReadData(string sqlSelect)
        {
            DataTable dt = new DataTable();
            OpenConncect();
            SqlDataAdapter sqldata = new SqlDataAdapter(sqlSelect, sqlConn);
            sqldata.Fill(dt);
            CloseConnect();
            sqldata.Dispose();
            return dt;
        }

    }
}
