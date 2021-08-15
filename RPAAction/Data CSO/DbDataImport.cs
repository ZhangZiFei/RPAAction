using System.Collections.Generic;
using System.Data.Common;
using System.Text;

namespace RPAAction.Data_CSO
{
    class DbDataImport : RPADataImport
    {
        public DbDataImport(DbConnection conn, string tableName, string brackets1 = "[", string brackets2 = "]")
        {
            this.conn = conn;
            this.tableName = tableName;
            this.brackets1 = brackets1;
            this.brackets2 = brackets2;
        }

        protected override void Close()
        {
            conn.Dispose();
        }

        protected override void CreateTable(DbDataReader r)
        {
            string type = "text";

            FieldCount = r.FieldCount;
            Values = new object[FieldCount];
            //InsterStrBuilder
            StringBuilder ISBer = new StringBuilder($@"INSERT INTO {brackets1}{tableName}{brackets2} (");
            for (int i = 0; i < r.FieldCount; i++)
            {
                ISBer.Append(brackets1);
                ISBer.Append(r.GetName(i));
                ISBer.Append(brackets2);
                ISBer.Append(",");
            }
            ISBer.Remove(ISBer.Length - 1, 1);
            ISBer.Append(") VALUES (");
            for (int i = 0; i < r.FieldCount; i++)
            {
                ISBer.Append("{" + i + "},");
            }
            ISBer.Remove(ISBer.Length - 1, 1);
            ISBer.Append(")");
            InsterStr = ISBer.ToString();

            //FieldMap
            for (int i = 0; i < r.FieldCount; i++)
            {
                FieldMap.Add(r.GetName(i), i);
            }

            var cmd = conn.CreateCommand();
            try
            {
                cmd.CommandText = GetCreateTableString(r, type, brackets1, brackets2);
                cmd.ExecuteNonQuery();
            }
            finally
            {
                cmd.Dispose();
            }
        }

        protected override void SetValue(string field, object value)
        {
            if (value == null)
            {
                Values[FieldMap[field]] = "NULL";
            }
            else
            {
                Values[FieldMap[field]] = "'" + value.ToString().Replace("'", "''") + "'";
            }
        }

        protected override void UpdataRow()
        {
            var cmd = conn.CreateCommand();
            cmd.CommandText = string.Format(InsterStr, Values);
            cmd.ExecuteNonQuery();
            Values = new object[FieldCount];
        }

        private readonly DbConnection conn;

        private readonly string brackets1;
        private readonly string brackets2;

        private string InsterStr = "";
        private int FieldCount;
        private object[] Values;
        private readonly Dictionary<string, int> FieldMap = new Dictionary<string, int>();
    }
}
