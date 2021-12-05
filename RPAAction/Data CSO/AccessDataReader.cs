using RPAAction.Access_CSO;
using Microsoft.Office.Interop.Access.Dao;
using System;

namespace RPAAction.Data_CSO
{
    public class AccessDataReader : RPADataReader
    {
        public readonly int Count;

        public AccessDataReader(string accessPath, string SQL)
        {
            accessInfo = new Internal_AccessInfo(accessPath);
            rd = accessInfo.Db.OpenRecordset(SQL);
            Count = rd.RecordCount;
        }

        public override int FieldCount => rd.Fields.Count;

        public override bool IsClosed => true;

        public override bool HasRows
        {
            ///////待测试
            get
            {
                try
                {
                    _ = rd.Fields[0].Value;
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        public override void Close()
        {
            accessInfo.Close();
        }

        public override string GetName(int ordinal)
        {
            return rd.Fields[ordinal].Name;
        }

        public override object GetValue(int ordinal)
        {
            return rd.Fields[ordinal].Value;
        }

        public override bool Read()
        {
            if (startRead)
            {
                rd.MoveNext();
            }
            else
            {
                startRead = true;
            }
            return !rd.EOF;
        }

        private readonly Internal_AccessInfo accessInfo;

        private readonly Recordset rd;

        private bool startRead = false;
    }
}
