using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction.Access_CSO
{
    public class Sql_GetAc : AccessAction
    {
        public decimal Count => count;

        public Sql_GetAc(string accessPath, string SQL)
            : base(accessPath)
        {
            this.SQL = SQL;
            Run();
        }

        public Base.RPAAction Out(out decimal Count)
        {
            Count = count;
            return this;
        }

        protected override void Action()
        {
            base.Action();

            db.Execute(SQL);
            count = db.RecordsAffected;
        }

        private readonly string SQL;

        private decimal count = 0;
    }
}
