﻿using RPAAction.Base;
using System;
using System.Data.SqlClient;

namespace RPAAction.Data_CSO
{
    public class SQLServerDataImport : RPADataImport
    {
        public SQLServerDataImport(string DataSource, string DataBase, string user, string pwd, string table)
        {
            this.connStr = string.Format(@"Data Source={0};Initial Catalog={1};User ID={2};Pwd={3};", DataSource, DataBase, user, pwd);
            conn = new SqlConnection(connStr);
            conn.Open();
            this.tableName = table;
        }

        public SQLServerDataImport(string connStr, string table)
        {
            this.connStr = connStr;
            conn = new SqlConnection(connStr);
            conn.Open();
            this.tableName = table;
        }

        public SQLServerDataImport(SqlConnection conn, string table)
        {
            this.conn = conn;
            this.tableName = table;
        }

        public override void ImportFrom(RPADataReader reader)
        {
            try
            {
                CreateTable(reader);
            }
            catch (Exception e)
            {
                if (e is ActionException)
                    throw e;
            }

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
            {
                bulkCopy.DestinationTableName = tableName;
                bulkCopy.WriteToServer(reader);
            }
        }

        public override void Dispose()
        {
            if (connStr == null)
            {
                conn.Dispose();
            }
        }

        protected override void setValue(string field, object value)
        {
            throw new NotImplementedException();
        }

        protected override void updataRow()
        {
            throw new NotImplementedException();
        }

        protected override void CreateTable(RPADataReader r)
        {
            string sql = GetCreateTableString(r, "text");
            var cmd = new SqlCommand(sql.ToString(), conn);
            cmd.ExecuteNonQuery();
        }

        private string connStr = null;
        private SqlConnection conn;
    }
}