using RPAAction.Base;
using System.IO;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 工作簿-另存为
    /// 如果 newWbPath 存在文件将抛出异常
    /// </summary>
    public class Workbook_SaveAs : ExcelAction
    {
        public Workbook_SaveAs(string wbPath = null, string newWbPath = null)
            : base(wbPath)
        {
            this.newWbPath = Path.GetFullPath(newWbPath);
            Run();
        }

        protected override void Action()
        {
            base.Action();

            if (File.Exists(newWbPath))
            {
                throw new ActionException($"文件({newWbPath})已经存在");
            }

            Directory.CreateDirectory(Path.GetDirectoryName(newWbPath));
            wb.SaveAs(newWbPath, GetXlFileFormatByWbPath(newWbPath));
        }

        private string newWbPath = null;
    }
}
