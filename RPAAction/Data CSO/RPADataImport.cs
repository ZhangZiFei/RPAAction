using RPAAction.Base;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction.Data_CSO
{
    /// <summary>
    /// RPA数据导入
    /// </summary>
    public abstract class RPADataImport : IDisposable
    {
        /// <summary>
        /// 数据导入,然后释放依赖
        /// </summary>
        /// <param name="i"></param>
        /// <param name="r"></param>
        public static void ImportDispose(RPADataImport i, RPADataReader r)
        {
            using (i)
            {
                using (r)
                {
                    i.ImportFrom(r);
                }
            }
        }

        /// <summary>
        /// 数据导入,然后释放依赖(异步)
        /// </summary>
        /// <param name="i"></param>
        /// <param name="r"></param>
        /// <returns></returns> 
        public static async Task ImportDisposeAsync(RPADataImport i, RPADataReader r)
        {
            await Task.Run(() => {
                ImportDispose(i, r);
            });
        }

        public abstract void Dispose();

        public virtual void ImportFrom(RPADataReader reader)
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

            int count = reader.FieldCount;
            while (reader.Read())
            {
                for (int i = 0; i < count; i++)
                {
                    setValue(reader.GetName(i), reader.GetValue(i));
                }
                updataRow();
            }
        }

        public virtual async Task ImportFromAsync(RPADataReader reader)
        {
            await Task.Run(()=> {
                ImportFrom(reader);
            });
        }

        protected string tableName;

        protected abstract void setValue(string field, object value);
        protected abstract void updataRow();
        protected abstract void CreateTable(RPADataReader r);

        protected string GetCreateTableString(RPADataReader r, string type)
        {
            StringBuilder sql = new StringBuilder("CREATE TABLE ");
            sql.Append(tableName);
            sql.Append("(");
            for (int i = 0; i < r.FieldCount; i++)
            {
                sql.Append("[");
                sql.Append(r.GetName(i));
                sql.Append("] ");
                sql.Append(type);
                sql.Append(",");
            }
            sql.Remove(sql.Length - 1, 1);
            sql.Append(")");
            return sql.ToString();
        }
    }
}
