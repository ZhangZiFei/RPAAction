using Microsoft.Office.Interop.Excel;

namespace RPAAction.Excel_CSO
{
    public static class WorkbookExpand
    {
        /// <summary>
        /// 处理<see cref="_Workbook"/>以适应自动化操作
        /// </summary>
        public static _Workbook ChangeForRPA(this _Workbook wb)
        {
            wb.CheckCompatibility = false;//控制兼容性检查器运行自动保存工作簿时。 为可读/写属性。
            wb.UpdateLinks = XlUpdateLinks.xlUpdateLinksNever;//禁止更新链接
            return wb;
        }
    }
}
