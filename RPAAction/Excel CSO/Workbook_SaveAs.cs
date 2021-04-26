using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RPAAction.Base;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 工作簿-另存为
    /// 如果 newWbPath 存在文件将抛出异常
    /// </summary>
    class Workbook_SaveAs : ExcelAction
    {
        public Workbook_SaveAs(string wbPath = null, string newWbPath = null)
            : base(wbPath)
        {
            this.newWbPath = Path.GetFullPath(newWbPath);
            Run();
        }

        protected override void action()
        {
            base.action();

            if (File.Exists(newWbPath))
            {
                throw new ActionException(string.Format("文件({0})已经存在", newWbPath));
            }

            Directory.CreateDirectory(Path.GetDirectoryName(newWbPath));
            wb.SaveAs(newWbPath, getXlFileFormatByWbPath(newWbPath));
        }

        private string newWbPath = null;
    }
}
