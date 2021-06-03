namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 工作簿-创建工作表
    /// </summary>
    public class Workbook_CreateWorksheet : ExcelAction
    {
        public new string wsName = null;

        ///<param name="wsName">新工作表的名称,默认由Excel进程自动设置</param>
        /// <param name="position">新工作表的位置,默认0,正常范围是1到工作表的最大数量,正常范围之外视为最后一个位置</param>
        /// <param name="before">提供一个工作表名称,新工作表创建在该工作表之前,如果提供此参数将无视参数"position"</param>
        /// <param name="after">提供一个工作表名称,新工作表创建在该工作表之后,如果提供此参数将无视参数"position"和"before"</param>
        public Workbook_CreateWorksheet(string wbPath = null, string wsName = null, decimal position = 0, string before = null, string after = null)
            : base(wbPath)
        {
            this.wsName = wsName;
            this.position = position;
            this.before = before;
            this.after = after;
            Run();
        }

        protected override void action()
        {
            base.action();
            //Create Worksheet
            if (CheckString(after))
            {
                ws = wb.Worksheets.Add(After: wb.Worksheets[after]);
            }
            else if (CheckString(before))
            {
                ws = wb.Worksheets.Add(Before: wb.Worksheets[before]);
            }
            else if (position >= 1 && position <= wb.Sheets.Count)
            {
                ws = wb.Worksheets.Add(Before: wb.Worksheets[position]);
            }
            else
            {
                ws = wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            }
            //wsName
            if (CheckString(wsName))
            {
                ws.Name = wsName;
            }
            else
            {
                wsName = ws.Name;
            }
            base.wsName = wsName;
        }

        private decimal position = 0;
        private string before = null;
        private string after = null;
    }
}
