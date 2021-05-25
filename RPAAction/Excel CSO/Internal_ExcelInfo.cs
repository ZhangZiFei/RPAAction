using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 内部-ExcelInfo
    /// 自动打开Excel并且获取想相关COM对象和相关信息
    /// </summary>
    class Internal_ExcelInfo : ExcelAction
    {
        /// <param name="wbPath">工作簿路径, 如果为空视为获取活动工作簿</param>
        /// <param name="wsName">工作表名称, 如果为空视为获取活动工作表</param>
        public Internal_ExcelInfo(string wbPath = null, string wsName = null, string range = null)
            : base(wbPath, wsName, range)
        {
            Run();
        }

        public new _Application app => ExcelAction.app;

        public new _Workbook wb => base.wb;

        public new _Worksheet ws => base.ws;

        public new Range R => base.R;

        /// <summary>
        /// 工作簿路径
        /// </summary>
        public new string wbPath => base.wbPath;

        /// <summary>
        /// 工作簿文件名(带后缀)
        /// </summary>
        public new string wbFileName => wbFileName;

        /// <summary>
        /// 工作表名称
        /// </summary>
        public new string wsName => wsName;

        /// <summary>
        /// 单元格名称
        /// </summary>
        public new string range => base.range;

        /// <summary>
        /// <see cref="ExcelAction.app"/>是否由当前的Action打开
        /// </summary>
        public new bool isOpenApp => base.isOpenApp;

        /// <summary>
        /// <see cref="wb"/>是否由当前Action打开
        /// </summary>
        public new bool isOpenWorkbook => base.isOpenWorkbook;
    }
}
