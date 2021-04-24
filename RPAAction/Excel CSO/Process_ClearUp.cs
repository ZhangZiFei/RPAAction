using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 进程-清理
    /// 处理<see cref="_Application"/>以适应用户操作
    /// </summary>
    class Process_ClearUp : ExcelAction
    {
        public Process_ClearUp()
        {
            Run();
        }

        protected override void action()
        {
            if(AttachApp() !=  null)
            {
                ChangeAppForUser(app);
            }
        }
    }
}
