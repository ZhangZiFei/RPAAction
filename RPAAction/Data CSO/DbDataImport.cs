using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction.Data_CSO
{
    class DbDataImport : RPADataImport
    {
        public DbDataImport(string sql, DbConnection conn)
        {
            this.conn = conn;
            cmd = conn.CreateCommand();
            cmd.CommandText = sql;
        }

        public override void Dispose()
        {

        }

        protected override void CreateTable(DbDataReader r)
        {
            string sql = GetCreateTableString(r, "text");
            var cmd = conn.CreateCommand();
            try
            {
                cmd.CommandText = sql;
                cmd.ExecuteNonQuery();
            }
            finally
            {
                cmd.Dispose();
            }
        }

        protected override void SetValue(string field, object value)
        {
            cmd.Parameters.Add(value);
        }

        protected override void UpdataRow()
        {
            throw new NotImplementedException();
        }

        private readonly DbConnection conn;

        private readonly DbCommand cmd;
    }
}
