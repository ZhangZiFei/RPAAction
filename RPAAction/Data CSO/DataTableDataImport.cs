using System.Data;

namespace RPAAction.Data_CSO
{
    public class DataTableDataImport : RPADataImport
    {
        public readonly DataTable table;

        public DataTableDataImport(DataTable table = null)
        {
            this.table = table == null ? new DataTable() : table;
        }

        public override void Dispose()
        {

        }

        protected override void CreateTable(RPADataReader r)
        {
            string rName;
            for (int i = 0; i < r.FieldCount; i++)
            {
                rName = r.GetName(i);
                if (! table.Columns.Contains(rName))
                {
                    table.Columns.Add(rName);
                }
            }
        }

        protected override void setValue(string field, object value)
        {
            if (table.Rows.Count <= writeRow)
                table.Rows.Add();
            table.Rows[writeRow][field] = value;
        }

        protected override void updataRow()
        {
            ++writeRow;
        }

        private int writeRow = 0;
    }
}
