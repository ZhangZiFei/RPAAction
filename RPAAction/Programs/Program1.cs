using System;
using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;
using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using System.Data.SqlClient;
using System.Data.Common;
using System.Runtime.Remoting;
using System.Collections;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using RPAAction.Data_CSO;

namespace RPAAction
{
    class Program
    {
        static void Main()
        {
            //DbConnection conn = new SqlConnection(@"Data Source=10.132.56.70,3000;Initial Catalog=zifeiTest;User ID=sa;Pwd=foxconn123!!;");
            //DbConnection conn = new OleDbConnection(@"Provider=MSOLEDBSQL;Data Source=10.132.56.70,3000;Initial Catalog=zifeiTest;User ID=sa;Pwd=foxconn123!!;");
            //DbConnection conn = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source =C:\Users\F1336747\Desktop\t.accdb");
            DbConnection conn = new OdbcConnection(@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\Users\F1336747\Desktop\t.accdb;");
            conn.Open();

            //RPADataImport.ImportDispose(
            //       new ExcelDataReader(@"C:\Users\F1336747\Desktop\vb.xlsm"),//, "", "used", 6),//
            //       //new SQLServerDataImport(conn, "a")
            //       new DataTableDataImport(a)
            //);
            //RPADataImport.ImportDispose(
            //       new DataTableDataReader(a),
            //       new SQLServerDataImport(conn, "a")
            //);
            //RPADataImport.ImportDispose(
            //       RPADataReader.GetDbDataReader(conn, "SELECT * FROM a"),
            //       new DataTableDataImport(a)
            //);
            //RPADataImport.ImportDispose(
            //    new ExcelDataReader(@"C:\Users\F1336747\Desktop\vb.xlsm", "a"),
            //    new ExcelDataImport(@"C:\Users\F1336747\Desktop\vb.xlsm", "b")
            //);
            RPADataImport.ImportDispose(
                new ExcelDataReader(@"C:\Users\F1336747\Desktop\vb.xlsm", "a"),
                new DbDataImport(conn, "b")
            );
            try
            {
                Excel_CSO.ExcelAction.ChangeAppForUser(Excel_CSO.ExcelAction.AttachApp());
            }
            catch { }
        }

        static void AccessTest()
        {
            Application app = new Application
            {
                Visible = true,
                UserControl = true
            };
            app.OpenCurrentDatabase(@"F:\testsrc\a.accdb");
            Database db = app.CurrentDb();
            Recordset rd = db.OpenRecordset("a");

            int t1 = Environment.TickCount;

            for (int i = 0; i < 100000; ++i)
            {
                rd.AddNew();
                rd.Fields["a"].Value = i;
                rd.Fields["b"].Value = "哈哈哈哈哈啊哈";
                rd.Fields["c"].Value = "233333333333333333333333333333";
                rd.Update();
                if (i % 1000 == 0)
                {
                    Console.WriteLine(i);
                }
            }

            Console.WriteLine("用时：" + (t1 - Environment.TickCount));
            Console.ReadLine();

            app.Quit();
        }
    }

    class RPADataTable : DataTable
    {

    }
}
