using RPAAction.Base;
using System.IO;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 进程-创建工作簿
    /// 如果创建的文件已经存在则抛出异常
    /// </summary>
    public class Process_CreateWorkbook : ExcelAction
    {
        public Process_CreateWorkbook(string wbPath)
        {
            this.WbPath = Path.GetFullPath(wbPath);
            Run();
        }

        protected override void Action()
        {
            AttachOrOpenApp();

            if (File.Exists(WbPath))
            {
                throw new ActionException($"文件({WbPath})已经存在");
            }

            Directory.CreateDirectory(Path.GetDirectoryName(WbPath));
            Wb = App.Workbooks.Add();
            Wb.SaveAs(WbPath, GetXlFileFormatByWbPath(WbPath));
        }
    }
}

