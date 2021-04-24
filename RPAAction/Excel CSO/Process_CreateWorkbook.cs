using RPAAction.Base;
using System.IO;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 进程-创建工作簿
    /// </summary>
    class Process_CreateWorkbook : ExcelAction
    {
        public Process_CreateWorkbook(string wbPath)
        {
            this.wbPath = wbPath;
            Run();
        }

        protected override void action()
        {
            AttachOrOpenApp();

            if (!CheckString(wbPath))
            {
                throw new ActionException("请指定参数\"wbPath\"");
            }

            //检查目录以及路径
            string dir = Path.GetDirectoryName(wbPath);
            if (Directory.Exists(dir))
            {
                if (File.Exists(wbPath))
                {
                    throw new ActionException(string.Format("文件({0})已经存在", wbPath));
                }
            }
            else
            {
                Directory.CreateDirectory(dir);
            }

            wb = app.Workbooks.Add();
            ws.SaveAs(wbPath);
        }
    }
}

