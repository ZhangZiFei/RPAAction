using Microsoft.Office.Interop.Excel;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 进程-关闭
    /// 自动当前用户下的所有Excel进程
    /// </summary>
    public class Process_Close : ExcelAction
    {
        public Process_Close()
        {
            Run();
        }

        protected override void Action()
        {
            if (!App.Check())
            {
                App = AttachApp();
            }
            while (App != null)
            {
                if (App.Check())
                {
                    //關閉應用和工作簿
                    foreach (_Workbook item in App.Workbooks)
                    {
                        item.Close(false);
                    }
                }
                App.Kill();
                App = AttachApp();
            }
        }
    }
}
