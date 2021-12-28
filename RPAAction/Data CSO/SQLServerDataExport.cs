using RPAAction.Base;
using System;
using System.Data.Common;
using System.Data.SqlClient;

namespace RPAAction.Data_CSO
{
    public class SQLServerDataExport : RPADataExport
    {
        /// <param name="DataSource"></param>
        /// <param name="DataBase"></param>
        /// <param name="user"></param>
        /// <param name="pwd"></param>
        /// <param name="table"></param>
        /// <param name="appand">是否附加数据默认true,否则清空表</param>
        /// <param name="bulkCopyTimeout"></param>
        public SQLServerDataExport(string DataSource, string DataBase, string user, string pwd, string table, bool appand = true, int bulkCopyTimeout = 600)
        {
            connStr = $@"Data Source={DataSource};Initial Catalog={DataBase};User ID={user};Pwd={pwd};";
            conn = new SqlConnection(connStr);
            conn.Open();
            tableName = table;
            BulkCopyTimeout = bulkCopyTimeout;
            this.appand = appand;
        }

        public SQLServerDataExport(string connStr, string table)
        {
            this.connStr = connStr;
            conn = new SqlConnection(connStr);
            conn.Open();
            tableName = table;
        }

        public SQLServerDataExport(SqlConnection conn, string table)
        {
            this.conn = conn;
            tableName = table;
        }

        public override void ImportFrom(DbDataReader reader)
        {
            try
            {
                if (reader.HasRows)
                {
                    CreateTable(reader);
                }
                else
                {
                    return;
                }
            }
            catch (Exception e)
            {
                if (e is ActionException)
                    throw e;
            }
            if (!appand)
            {
                var cmd = new SqlCommand($"TRUNCATE TABLE {tableName};", conn);
                cmd.ExecuteNonQuery();
            }
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
            {
                bulkCopy.BulkCopyTimeout = BulkCopyTimeout;
                bulkCopy.DestinationTableName = tableName;
                bulkCopy.WriteToServer(reader);
            }
        }

        protected override void Close()
        {
            if (connStr == null)
            {
                conn.Dispose();
            }
        }

        protected override void SetValue(string field, object value)
        {
            throw new NotImplementedException();
        }

        protected override void UpdataRow()
        {
            throw new NotImplementedException();
        }

        protected override void CreateTable(DbDataReader r)
        {
            string sql = GetCreateTableString(r, "text");
            var cmd = new SqlCommand(sql.ToString(), conn);
            cmd.ExecuteNonQuery();
        }

        private readonly string connStr = null;
        private readonly SqlConnection conn;
        private readonly int BulkCopyTimeout;
        private readonly bool appand;
    }
}
