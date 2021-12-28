using RPAAction.Data_CSO;
using System;
using RPAAction.Excel_CSO;

namespace RPAAction
{
    class Program
    {
        static void Main(params string[] p)
        {
            if (p.Length > 0)
            {
                if (p[0].Equals("run"))
                {
                    Run.Run_(p);
                    return;
                }
            }
            new RPAAction();//F();//

            /////////////////////////////////////////////////////////////////////////////////////////
            //Vba();
            //EXE.DataAction.SqlServerToExcel(p[0], p[1], p[2], p[3], p[4], p[5], p[6]);
            //Console.InputEncoding = System.Text.Encoding.UTF8;
            //Console.OutputEncoding = System.Text.Encoding.UTF8;
            //DateTime beforDT = System.DateTime.Now;

            //F();

            //DateTime afterDT = System.DateTime.Now;
            //TimeSpan ts = afterDT.Subtract(beforDT);
            //Console.WriteLine("总共花费{0}s.", ts.TotalSeconds);
            //Console.ReadLine();
        }

        public static void Access_CSO_Test()
        {
            RPADataExport.ImportDispose(
                new ExcelDataReader(
                    @"C:\Users\zhang\Desktop\a.xlsx", "d"
                ),
                //new AccessDataImport(
                //    @"C:\Users\zhang\Desktop\a.accdb",
                //    "d"
                //)
                //请考虑使用 OLEDB 而不是 ODBC 
                //new DbDataImport(
                //    new System.Data.Odbc.OdbcConnection(@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\Users\zhang\Desktop\a.accdb;"),
                //    "d"
                //)
                new DbDataExport(
                    new System.Data.OleDb.OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\zhang\Desktop\a.accdb;"),
                    "d"
                )
            );
            //AccessAction.ClearUp();
            new Process_Close();
        }

        public static void F()
        {
        }

        public static void Vba()
        {
            var a= new HighLevel_RunFunction(@"C:\Users\zhang\Desktop\a.xlsx", "", System.IO.File.ReadAllText(@"a.vbs"), "f", "233");
            Console.WriteLine(a.Ret);
            new Process_Close();
        }
    }
}