using RPAAction.Data_CSO;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction.EXE
{
    static public class DataAction
    {
        public static void SqlServerToExcel(string DataSource, string DataBase, string user, string pwd, string SQL, string ExcelPath, string sheet)
        {
            //EXE.DataAction.SqlServerToExcel(p[0], p[1], p[2], p[3], p[4], p[5], p[6]);
            string connStr = $@"Data Source={DataSource};Initial Catalog={DataBase};User ID={user};Pwd={pwd};";
            using (SqlConnection connection = new SqlConnection(connStr))
            {
                SqlCommand command = new SqlCommand(SQL, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                RPADataImport.ImportDispose(
                    reader,
                    new ExcelDataImport(ExcelPath, sheet, "A1", true)
                );
            }
        }
    }
}
