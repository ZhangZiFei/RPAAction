namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 工作簿-删除工作表
    /// </summary>
    class Workbook_DeleteWorksheet : ExcelAction
    {
        public Workbook_DeleteWorksheet(string wbPath = null, string wsName = null)
            : base(wbPath, wsName)
        {
            Run();
        }

        protected override void action()
        {
            base.action();
            ws.Delete();
        }
    }
}
