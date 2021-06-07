﻿using RPAAction.Base;
using System;
using System.Data.Common;
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
        public static void ImportDispose(DbDataReader r, RPADataImport i)
        {
            using (r)
            {
                using (i)
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
        public static async Task ImportDisposeAsync(DbDataReader r, RPADataImport i)
        {
            await Task.Run(() => {
                ImportDispose(r, i);
            });
        }

        public abstract void Dispose();

        public virtual void ImportFrom(DbDataReader reader)
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
                    SetValue(reader.GetName(i), reader.GetValue(i));
                }
                UpdataRow();
            }
        }

        public virtual async Task ImportFromAsync(DbDataReader reader)
        {
            await Task.Run(()=> {
                ImportFrom(reader);
            });
        }

        protected string tableName;

        protected abstract void SetValue(string field, object value);
        protected abstract void UpdataRow();
        protected abstract void CreateTable(DbDataReader r);

        protected string GetCreateTableString(DbDataReader r, string type)
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
