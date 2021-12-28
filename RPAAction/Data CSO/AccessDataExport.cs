using System;
using System.Data.Common;
using RPAAction.Access_CSO;
using Microsoft.Office.Interop.Access.Dao;

namespace RPAAction.Data_CSO
{
    public class AccessDataExport : RPADataExport
    {
        public AccessDataExport(string accessPath, string table)
        {
            accessInfo = new Internal_AccessInfo(accessPath);
            db = accessInfo.App.CurrentDb();
            tableName = table;
        }

        protected override void Close()
        {
            accessInfo.Close();
        }

        protected override void CreateTable(DbDataReader r)
        {
            Exception sqlE = null;
            try
            {
                string sql = GetCreateTableString(r, "text");
                db.Execute(sql);
            }
            catch (Exception e)
            {
                sqlE = e;
                throw;
            }
            finally
            {
                try
                {
                    rd = db.OpenRecordset(tableName);
                    rd.AddNew();
                }
                catch (Exception)
                {
                    throw new Base.ActionException($"创建表[{tableName}]失败,详情如下:\n{sqlE ?? new Exception("")}");
                }
            }
        }

        protected override void SetValue(string field, object value)
        {
            rd.Fields[field].Value = value;
        }

        protected override void UpdataRow()
        {
            rd.Update();
            rd.AddNew();
        }

        private readonly Internal_AccessInfo accessInfo;

        private readonly Database db;

        private Recordset rd;
    }
}
