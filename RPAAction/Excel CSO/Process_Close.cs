using Microsoft.Office.Interop.Excel;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 进程-关闭
    /// 自动当前用户下的所有Excel进程
    /// </summary>
    class Process_Close : ExcelAction
    {
        public Process_Close()
        {
            Run();
        }

        protected override void action()
        {
            if (!CheckApp(app))
            {
                app = AttachApp();
            }
            while (app != null)
            {
                if (CheckApp(app))
                {
                    //關閉應用和工作簿
                    foreach (_Workbook item in app.Workbooks)
                    {
                        item.Close(false);
                    }
                }
                KillApp(app);
                app = AttachApp();
            }
        }
    }
}
