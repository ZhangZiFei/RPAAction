using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 进程-关闭工作簿
    /// </summary>
    public class Process_CloseWorkbook : ExcelAction
    {
        public Process_CloseWorkbook(string wbPath = null, bool isSave  = false)
            : base(wbPath)
        {
            this.isSave = isSave;
            Run();
        }

        protected override void action()
        {
            if (CheckString(wbPath))
            {
                wb = AttachWorkbook(wbPath);
                if (wb != null)
                {
                    wb.Close(isSave);
                }
            }
            else
            {
                base.action();
                wb.Close(isSave);
            }
        }

        private readonly bool isSave;
    }
}
