using Microsoft.Office.Interop.Excel;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 工作簿-获取工作表列表
    /// </summary>
    public class Workbook_GetWorksheetList : ExcelAction
    {
        public System.Data.DataTable table = null;

        public Workbook_GetWorksheetList(string wbPath = null)
            : base(wbPath)
        {
            Run();
        }

        protected override void Action()
        {
            base.Action();
            InitTable();
            _Worksheet _ws;
            foreach (object ws in Wb.Worksheets)
            {
                _ws = ws as _Worksheet;
                if (_ws != null)
                {
                    table.Rows.Add(_ws.Name);
                }
            }

        }

        private void InitTable()
        {
            table = new System.Data.DataTable();
            table.Columns.Add("Name");
        }
    }
}
