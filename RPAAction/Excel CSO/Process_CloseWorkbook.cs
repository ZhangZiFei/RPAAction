namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 进程-关闭工作簿
    /// </summary>
    public class Process_CloseWorkbook : ExcelAction
    {
        public Process_CloseWorkbook(string wbPath = null, bool isSave = false)
            : base(wbPath)
        {
            this.isSave = isSave;
            Run();
        }

        protected override void Action()
        {
            if (WbPath.CheckNoVoid())
            {
                Wb = AttachWorkbook(WbPath);
                if (Wb != null)
                {
                    Wb.Close(isSave);
                }
            }
            else
            {
                base.Action();
                Wb.Close(isSave);
            }
        }

        private readonly bool isSave;
    }
}
