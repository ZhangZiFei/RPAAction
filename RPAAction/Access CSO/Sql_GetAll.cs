using RPAAction.Data_CSO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction.Access_CSO
{
    public class Sql_GetAll : AccessAction
    {
        public DataTable Result => result;

        public decimal Count => count;

        public Sql_GetAll Out(out DataTable Result, out decimal Count)
        {
            Result = result;
            Count = count;
            return this;
        }

        public Sql_GetAll(string accessPath, string SQL)
            : base(accessPath)
        {
            this.SQL = SQL;
            Run();
        }

        protected override void Action()
        {
            AccessDataReader accessDataReader = new AccessDataReader(accessPath, SQL) ;
            count = accessDataReader.Count;
            RPADataExport.ImportDispose(accessDataReader, new DataTableDataExport(result));
        }

        private readonly DataTable result = new DataTable();

        private decimal count = 0;

        private readonly string SQL;
    }
}
