using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction.Data_CSO
{
    public class AccessToExcel : DataAction
    {
        public AccessToExcel(string AccessPath, string AccessTableName, string ExcelPAth, string ExcelSheetName)
        {
            this.AccessPath = System.IO.Path.GetFileName(AccessPath);
            this.AccessTableName = AccessTableName;
            this.ExcelPAth = System.IO.Path.GetFileName(ExcelPAth);
            this.ExcelSheetName = ExcelSheetName;
        }

        protected override void action()
        {

        }

        private string AccessPath;
        private string AccessTableName;
        private string ExcelPAth;
        private string ExcelSheetName;
    }
}
