using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction.CSRFC
{
    class IRfcTableRPADataReader : Data_CSO.RPADataReader
    {
        public IRfcTableRPADataReader(IRfcTable rfcTable)
        {
            IRfcTable = rfcTable;
            enumer =  IRfcTable.GetEnumerator();
        }

        public override int FieldCount => IRfcTable.ElementCount;

        public override bool HasRows => IRfcTable.Count > 0;

        public override void Close()
        {
        }

        public override string GetName(int ordinal)
        {
            return IRfcTable.GetElementMetadata(ordinal).Name;
        }

        public override object GetValue(int ordinal)
        {
            object o = enumer.Current.GetString(ordinal);
            //if (o is byte[])
            //{
            //    o = Encoding.UTF8.GetString((byte[])o);
            //}
            return o;
        }

        public override bool Read()
        {
            return enumer.MoveNext();
        }

        private readonly IRfcTable IRfcTable;
        private readonly IEnumerator<IRfcStructure> enumer;
    }
}
