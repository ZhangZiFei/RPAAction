using Microsoft.Office.Interop.Excel;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 进程-获取工作簿列表
    /// </summary>
    public class Process_GetWorkbookList : ExcelAction
    {
        public System.Data.DataTable table = null;

        public Process_GetWorkbookList()
        {
            Run();
        }

        protected override void Action()
        {
            AttachApp();
            if (App != null)
            {
                InitTable();
                foreach (_Workbook wb in App.Workbooks)
                {
                    table.Rows.Add(wb.Name, wb.FullName);
                }
            }
        }

        private void InitTable()
        {
            table = new System.Data.DataTable();
            table.Columns.Add("Name");
            table.Columns.Add("FullName");
        }
    }
}
