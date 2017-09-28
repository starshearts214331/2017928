using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wk2_2
{
    using System.Data;
    using System.Data.OleDb;
    using System.Xml.Serialization;
    class Program
    {
        static void Main(string[] args)
        {
            string 連線字串 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=BugTypes.MDB";
            string 查詢字串 = "SELECT * FROM Categories";
            DataTable TB = 連結資料庫類別.取得資料表類別方法(連線字串,查詢字串);
            Console.WriteLine(TB.Rows.Count.ToString());
            for (int i = 0; i < TB.Rows.Count; i++)
            {
                Console.WriteLine(TB.Rows[i].ItemArray[0].ToString() + "," + TB.Rows[i].ItemArray[1].ToString());
            }
        }
        class 連結資料庫類別
        {
            public static DataTable 取得資料表類別方法(string 連線字串, string 查詢字串)
            {
                DataTable DTB = new DataTable();
                OleDbConnection myAccessConn = myAccessConn = new OleDbConnection(連線字串);
                OleDbCommand myAccessCommand = new OleDbCommand(查詢字串, myAccessConn);
                myAccessConn.Open();
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(DTB);
                return DTB;
            }
            public DataTable 取得資料表物件方法(string 連線字串, string 查詢字串)
            {
                DataTable DTB = new DataTable();
                OleDbConnection myAccessConn = myAccessConn = new OleDbConnection(連線字串);
                OleDbCommand myAccessCommand = new OleDbCommand(查詢字串, myAccessConn);
                myAccessConn.Open();
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
                myDataAdapter.Fill(DTB);
                return DTB;
            }
        }
    }

  
}

