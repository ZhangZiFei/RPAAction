using System;
using System.Data.Common;
using System.Collections;
using System.Data;

namespace RPAAction.Data_CSO
{
    /// <summary>
    /// RPA数据流式读取
    /// </summary>
    public abstract class RPADataReader : DbDataReader
    {
        public static DbDataReader GetDbDataReader(DbConnection conn, string sql)
        {
            DbCommand cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            return cmd.ExecuteReader();
        }

        [Obsolete]
        public override object this[int ordinal] => throw new NotImplementedException();

        [Obsolete]
        public override object this[string name] => throw new NotImplementedException();

        [Obsolete]
        public override int Depth => throw new NotImplementedException();

        //public override int FieldCount => throw new NotImplementedException();

        [Obsolete]
        public override bool HasRows => throw new NotImplementedException();

        //public override bool IsClosed => throw new NotImplementedException();

        [Obsolete]
        public override int RecordsAffected => throw new NotImplementedException();

        //public override void Close()
        //{
        //    throw new NotImplementedException();
        //}

        [Obsolete]
        public override bool GetBoolean(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override byte GetByte(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override long GetBytes(int ordinal, long dataOffset, byte[] buffer, int bufferOffset, int length)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override char GetChar(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override long GetChars(int ordinal, long dataOffset, char[] buffer, int bufferOffset, int length)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override string GetDataTypeName(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override DateTime GetDateTime(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override decimal GetDecimal(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override double GetDouble(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override IEnumerator GetEnumerator()
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override Type GetFieldType(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override float GetFloat(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override Guid GetGuid(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override short GetInt16(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override int GetInt32(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override long GetInt64(int ordinal)
        {
            throw new NotImplementedException();
        }

        //public override string GetName(int ordinal)
        //{
        //    throw new NotImplementedException();
        //}

        [Obsolete]
        public override int GetOrdinal(string name)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override DataTable GetSchemaTable()
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override string GetString(int ordinal)
        {
            throw new NotImplementedException();
        }

        //public override object GetValue(int ordinal)
        //{
        //    throw new NotImplementedException();
        //}

        [Obsolete]
        public override int GetValues(object[] values)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override bool IsDBNull(int ordinal)
        {
            throw new NotImplementedException();
        }

        [Obsolete]
        public override bool NextResult()
        {
            throw new NotImplementedException();
        }

        //public override bool Read()
        //{
        //    throw new NotImplementedException();
        //}
    }
}
