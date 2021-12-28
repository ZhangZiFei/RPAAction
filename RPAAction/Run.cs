using RPAAction.Data_CSO;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction
{
    class Run
    {
        internal static void Run_(string[] p)
        {
            if (p.Length > 1)
            {
                if (p[1].Equals("SqlServerToExcel"))
                {
                    SqlServerToExcel(p);
                }
                else if (p[1].Equals("SqlServerToTXT"))
                {
                    SqlServerToTXT(p);
                }
                else if (p[1].Equals("ExcelToSqlServer"))
                {
                    ExcelToSqlServer(p);
                }
            }
            else
            {
                Console.WriteLine("SqlServerToExcel");
                Console.WriteLine("SqlServerToTXT");
                Console.WriteLine("ExcelToSqlServer");
            }
        }

        static void SqlServerToExcel(string[] p)
        {
            if (p.Length < 3)
            {
                Console.WriteLine("DataSource\tip,prot");//2
                Console.WriteLine("DataBase\t数据库名称");//3
                Console.WriteLine("user\t\t用户");//4
                Console.WriteLine("pwd\t\t密码");//5
                Console.WriteLine("SQL");//6
                Console.WriteLine("ExcelPath");//7
                Console.WriteLine("sheet");//8
            }
            else
            {
                string connStr = $@"Data Source={p[2]};Initial Catalog={p[3]};User ID={p[4]};Pwd={p[5]};";
                using (SqlConnection connection = new SqlConnection(connStr))
                {
                    SqlCommand command = new SqlCommand(p[6], connection);
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    RPADataExport.ImportDispose(
                        reader,
                        new ExcelDataExport(p[7], p[8], "A1", true)
                    );
                }
            }
        }

        static void SqlServerToTXT(string[] p)
        {
            if (p.Length < 3)
            {
                Console.WriteLine("DataSource\tip,prot");//2
                Console.WriteLine("DataBase\t数据库名称");//3
                Console.WriteLine("user\t\t用户");//4
                Console.WriteLine("pwd\t\t密码");//5
                Console.WriteLine("SQL");//6
                Console.WriteLine("Path");//7
            }
            else
            {
                string connStr = $@"Data Source={p[2]};Initial Catalog={p[3]};User ID={p[4]};Pwd={p[5]};";
                using (SqlConnection connection = new SqlConnection(connStr))
                {
                    SqlCommand command = new SqlCommand(p[6], connection);
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    RPADataExport.ImportDispose(
                        reader,
                        new TXTDataExport(p[7])
                    );
                }
            }
        }

        static void ExcelToSqlServer(string[] p)
        {
            if (p.Length < 3)
            {
                Console.WriteLine("DataSource\tip,prot");//2
                Console.WriteLine("DataBase\t数据库名称");//3
                Console.WriteLine("user\t\t用户");//4
                Console.WriteLine("pwd\t\t密码");//5
                Console.WriteLine("Table");//6
                Console.WriteLine("ExcelPath");//7
                Console.WriteLine("sheet");//8
                Console.WriteLine("TRUNCATE TABLE\t如果不是0则清空数据表");//9
            }
            else
            {
                RPADataExport.ImportDispose(
                    new ExcelDataReader(p[7], p[8]),
                    new SQLServerDataExport(p[2], p[3], p[4], p[5], p[6], p[9].Equals("0"))//超时时间半小时
                );
            }
        }
    }
}
