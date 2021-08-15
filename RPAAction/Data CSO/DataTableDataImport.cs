using System.Data;
using System.Data.Common;

namespace RPAAction.Data_CSO
{
    public class DataTableDataImport : RPADataImport
    {
        public readonly DataTable table;

        public DataTableDataImport(DataTable table = null)
        {
            this.table = table ?? new DataTable();
        }

        protected override void Close()
        {

        }

        protected override void CreateTable(DbDataReader r)
        {
            string rName;
            for (int i = 0; i < r.FieldCount; i++)
            {
                rName = r.GetName(i);
                if (!table.Columns.Contains(rName))
                {
                    table.Columns.Add(rName);
                }
            }
        }

        protected override void SetValue(string field, object value)
        {
            if (table.Rows.Count <= writeRow)
                table.Rows.Add();
            table.Rows[writeRow][field] = value;
        }

        protected override void UpdataRow()
        {
            ++writeRow;
        }

        private int writeRow = 0;
    }
}
