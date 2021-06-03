using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction.Data_CSO
{
    class DataTableDataReader : RPADataReader
    {
        public DataTableDataReader(DataTable table)
        {
            this.table = table;
        }

        public override int FieldCount => table.Columns.Count;

        public override bool IsClosed => true;

        public override void Close()
        {

        }

        public override string GetName(int ordinal)
        {
            return table.Columns[ordinal].ColumnName;
        }

        public override object GetValue(int ordinal)
        {
            return table.Rows[readingRowIndex][ordinal];
        }

        public override bool Read()
        {
            return ++readingRowIndex < table.Rows.Count;
        }
        
        private int readingRowIndex = -1;
        private readonly DataTable table;
    }
}
